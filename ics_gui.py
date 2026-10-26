"""Preview-only Tk application. Workers never call Tk or mutate appointments."""

from dataclasses import dataclass, replace
from datetime import date, datetime, timezone
from queue import Empty, Queue
import threading
import tkinter as tk
from tkinter import messagebox, ttk
from uuid import UUID, uuid4
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from ics_model import OperationKind
from ics_outlook import inspect_calendar, pick_calendar
from ics_profiles import Destination, ImportProfile, load_profiles, save_profiles, update_profile
from ics_run import ImportCoordinator


class FormError(ValueError):
    pass


class MappingChangeRequired(FormError):
    pass


@dataclass(frozen=True)
class ProfileForm:
    profile_id: UUID
    name: str
    source_url: str
    destination: Destination | None
    timezone_name: str
    start_text: str
    end_text: str


def new_form(today=None):
    today = today or date.today()
    return ProfileForm(uuid4(), "New profile", "", None, "Europe/Berlin",
                       f"{today.year}-01-01", f"{today.year + 1}-01-01")


def build_profile(form, *, base=None, accept_mapping_change=False):
    if not form.name.strip():
        raise FormError("Enter a profile name.")
    if form.destination is None:
        raise FormError("Choose an Outlook calendar.")
    if not form.timezone_name.strip():
        raise FormError("Enter a timezone, for example Europe/Berlin.")
    try:
        start, end = date.fromisoformat(form.start_text), date.fromisoformat(form.end_text)
        if start.isoformat() != form.start_text or end.isoformat() != form.end_text or end <= start:
            raise ValueError
    except (ValueError, TypeError):
        raise FormError("Use YYYY-MM-DD dates, with Until later than From. Until is exclusive.") from None
    try:
        candidate = ImportProfile(form.profile_id, form.name.strip(), form.source_url.strip(),
                                  form.destination, form.timezone_name.strip(), start, end)
    except ValueError:
        raise FormError("Enter a valid HTTPS ICS URL without credentials or a fragment.") from None
    if base is None:
        return candidate
    if candidate.profile_id != base.profile_id:
        raise FormError("Select the profile again before editing it.")
    changed = candidate.source_url != base.source_url or (
        candidate.destination.store_id.casefold(), candidate.destination.folder_id.casefold()) != (
        base.destination.store_id.casefold(), base.destination.folder_id.casefold())
    if changed and not accept_mapping_change:
        raise MappingChangeRequired("Confirm the source or destination change.")
    mapped = update_profile(base, source_url=candidate.source_url, destination=candidate.destination,
                            confirm_source_change=accept_mapping_change,
                            confirm_destination_change=accept_mapping_change)
    return replace(mapped, name=candidate.name, floating_timezone=candidate.floating_timezone,
                   window_start=start, window_end=end)


_ACTIONS = {OperationKind.CREATE: "New", OperationKind.UPDATE: "Update",
            OperationKind.UNCHANGED: "Unchanged", OperationKind.DELETE: "Delete",
            OperationKind.CONFLICT: "Conflict"}
_REASONS = {
    "Source present": "Present in source.",
    "Source projection differs": "Details or availability differ.",
    "Explicit source cancellation": "Cancelled in source.",
    "Source identity absent": "Missing from the current source snapshot.",
    "Source export coverage is unconfirmed": "The source's complete export period is unconfirmed.",
    "Source absence is outside confirmed export coverage": "Outside the confirmed source export period.",
    "Original slot and target start must both be in scope": "Original or current start is outside the selected period.",
    "Deletion preview requires explicit confirmation": "Deletion is not enabled in this preview.",
    "Empty snapshot requires separate confirmation": "Empty calendar download requires review.",
    "Large removal requires separate confirmation": "An unusually large number of removals requires review.",
    "Resolve all source/target conflicts before deleting": "Resolve the listed conflicts first.",
}


def preview_rows(preview):
    try:
        zone = ZoneInfo(preview.profile.floating_timezone)
        label = preview.profile.floating_timezone
    except (ZoneInfoNotFoundError, ValueError):
        zone, label = timezone.utc, "UTC (custom profile timezone is not available for display)"
    def stamp(value):
        return value.astimezone(zone).strftime("%Y-%m-%d %H:%M:%S") if isinstance(value, datetime) else value.isoformat()
    rows = []
    for operation in preview.plan.operations:
        fields = operation.source.fields if operation.source is not None else None
        if fields is None and operation.targets:
            fields = operation.targets[0].fields
        reason = ("The existing import cannot be matched safely." if operation.kind is OperationKind.CONFLICT
                  else _REASONS.get(operation.reason, "Review this proposed change."))
        rows.append((_ACTIONS[operation.kind], fields.subject if fields else "Event details unavailable",
                     stamp(fields.start) if fields else "", stamp(fields.end) if fields else "", reason))
    for candidate in preview.plan.pending_deletions:
        fields = candidate.target.fields
        reason = " ".join(_REASONS.get(text, "Requires review.") for text in candidate.blockers)
        rows.append(("Deletion candidate", fields.subject, stamp(fields.start), stamp(fields.end), reason))
    return tuple(rows), label


@dataclass(frozen=True)
class JobResult:
    kind: str
    generation: int
    value: object = None
    error_type: str | None = None


class ReadJobs:
    def __init__(self):
        self.busy = False
        self._queue = Queue()

    def start(self, kind, generation, operation):
        if self.busy:
            raise RuntimeError("A read operation is already running")
        self.busy = True
        def worker():
            try:
                result = JobResult(kind, generation, value=operation())
            except BaseException as error:
                error_type = getattr(error, "error_type", type(error).__name__)
                result = JobResult(kind, generation, error_type=error_type if isinstance(error_type, str) else type(error).__name__)
            self._queue.put(result)
        threading.Thread(target=worker, name="ics-preview-read", daemon=True).start()

    def poll(self):
        try:
            result = self._queue.get_nowait()
        except Empty:
            return None
        self.busy = False
        return result


def read_error_message(kind, error_type):
    if error_type == "RunBusyError":
        return "Another importer run is using this profile and calendar. Try again when it finishes."
    if error_type == "FeedDownloadError":
        return "Cannot download the calendar. Check the HTTPS URL and your network connection."
    if error_type == "IcsParseError":
        return "The source calendar is invalid or contains an unsupported event. No preview was produced."
    if error_type == "OutlookReadError":
        return "Cannot use this calendar. Open classic Outlook and choose an accessible calendar folder."
    return "The read operation could not be completed. No appointments were changed."


class ImporterWindow:
    def __init__(self, root, *, profiles_path=None, coordinator=None, picker=pick_calendar,
                 inspector=inspect_calendar, dialogs=messagebox):
        self.root, self.profiles_path, self.dialogs = root, profiles_path, dialogs
        self.coordinator = coordinator or ImportCoordinator()
        self.picker, self.inspector = picker, inspector
        self.jobs = ReadJobs()
        self.last_preview = None
        self._base = None
        self._generation = 0
        self._loading = self._closing = False
        self._alive = True
        self._profile_load_error = False
        try:
            self.profiles = list(load_profiles(profiles_path))
        except ValueError:
            self.profiles = []
            self._profile_load_error = True
        self.values = {name: tk.StringVar(root) for name in ("name", "url", "zone", "start", "end", "destination")}
        self.status = tk.StringVar(root, "Choose a source and destination, then build a preview.")
        self.calendar_hint = tk.StringVar(root, "No calendar selected.")
        self.summary = tk.StringVar(root, "No preview yet.")
        self.time_hint = tk.StringVar(root, "All-day end dates are exclusive.")
        self.selection_note = tk.StringVar(root, "Select a row to see its full review note.")
        self._build()
        for value in self.values.values():
            value.trace_add("write", self._changed)
        self._populate(self.profiles[0] if self.profiles else None)
        if self._profile_load_error:
            self.status.set("Profiles could not be loaded. Saving is disabled to preserve the existing file.")
        self.root.protocol("WM_DELETE_WINDOW", self.close)
        self.root.bind("<Configure>", self._resize)
        self._poll_id = self.root.after(80, self._poll)

    def _build(self):
        self.root.title("ICS Calendar Importer — Preview")
        self.root.geometry("1120x760")
        self.root.minsize(920, 600)
        frame = ttk.Frame(self.root, padding=16)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(9, weight=1)
        ttk.Label(frame, text="ICS Calendar Importer", font=("Segoe UI", 16, "bold")).grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(frame, text="Preview only — no appointments will be changed.").grid(row=1, column=0, columnspan=3, sticky="w", pady=(2, 14))
        toolbar = ttk.Frame(frame)
        toolbar.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(0, 12))
        toolbar.columnconfigure(1, weight=1)
        ttk.Label(toolbar, text="Profile").grid(row=0, column=0, padx=(0, 10))
        self.profile_combo = ttk.Combobox(toolbar, state="readonly")
        self.profile_combo.grid(row=0, column=1, sticky="ew")
        self.profile_combo.bind("<<ComboboxSelected>>", self._select_profile)
        self.new_button = ttk.Button(toolbar, text="New", command=lambda: self._populate(None))
        self.new_button.grid(row=0, column=2, padx=(8, 4))
        self.save_button = ttk.Button(toolbar, text="Save profile", command=self.save_profile)
        self.save_button.grid(row=0, column=3)
        if self._profile_load_error:
            ttk.Label(toolbar, text="Profiles file unreadable: saving is disabled. Repair the file and restart to save profiles.",
                      foreground="#9a2929", wraplength=800).grid(row=1, column=0, columnspan=4, sticky="w", pady=(6, 0))
        self.editable = []
        for row, label, name in ((3, "Profile name", "name"), (4, "Source (HTTPS ICS URL)", "url")):
            ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", padx=(0, 12), pady=5)
            entry = ttk.Entry(frame, textvariable=self.values[name])
            entry.grid(row=row, column=1, columnspan=2, sticky="ew", pady=5)
            self.editable.append(entry)
        ttk.Label(frame, text="Outlook destination").grid(row=5, column=0, sticky="w", padx=(0, 12), pady=5)
        ttk.Entry(frame, textvariable=self.values["destination"], state="readonly").grid(row=5, column=1, sticky="ew", pady=5)
        self.pick_button = ttk.Button(frame, text="Choose calendar…", command=self.choose_calendar)
        self.pick_button.grid(row=5, column=2, padx=(8, 0))
        self.hint_label = ttk.Label(frame, textvariable=self.calendar_hint, wraplength=840)
        self.hint_label.grid(row=6, column=0, columnspan=3, sticky="w", pady=(1, 10))
        dates = ttk.Frame(frame)
        dates.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(0, 12))
        for column, label, name, width in ((0, "From (YYYY-MM-DD)", "start", 13),
                                           (2, "Until (exclusive)", "end", 13),
                                           (4, "Timezone", "zone", 24)):
            ttk.Label(dates, text=label).grid(row=0, column=column, padx=(0, 7))
            entry = ttk.Entry(dates, textvariable=self.values[name], width=width)
            entry.grid(row=0, column=column + 1, padx=(0, 16))
            self.editable.append(entry)
        actions = ttk.Frame(frame)
        actions.grid(row=8, column=0, columnspan=3, sticky="ew", pady=(0, 8))
        self.preview_button = ttk.Button(actions, text="Build preview", command=self.build_preview)
        self.preview_button.pack(side="left")
        self.progress = ttk.Progressbar(actions, mode="indeterminate", length=150)
        self.progress.pack(side="left", padx=14)
        self.summary_label = ttk.Label(actions, textvariable=self.summary, wraplength=600)
        self.summary_label.pack(side="left")
        table = ttk.Frame(frame)
        table.grid(row=9, column=0, columnspan=3, sticky="nsew")
        table.rowconfigure(0, weight=1)
        table.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(table, columns=("action", "title", "start", "end", "reason"), show="headings", height=12)
        for name, label, width in (("action", "Action", 130), ("title", "Appointment", 260),
                                   ("start", "Start", 140), ("end", "End", 140), ("reason", "Review note", 410)):
            self.tree.heading(name, text=label)
            self.tree.column(name, width=width, minwidth=90, stretch=name in ("title", "reason"))
        self.tree.grid(row=0, column=0, sticky="nsew")
        self.tree.bind("<<TreeviewSelect>>", self._show_note)
        vertical = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal = ttk.Scrollbar(table, orient="horizontal", command=self.tree.xview)
        horizontal.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        self.time_label = ttk.Label(frame, textvariable=self.time_hint, wraplength=1000)
        self.time_label.grid(row=10, column=0, columnspan=3, sticky="w", pady=(8, 4))
        self.note_label = ttk.Label(frame, textvariable=self.selection_note, wraplength=1000)
        self.note_label.grid(row=11, column=0, columnspan=3, sticky="w", pady=(0, 8))
        self.status_label = ttk.Label(frame, textvariable=self.status, wraplength=1000)
        self.status_label.grid(row=12, column=0, columnspan=3, sticky="w")

    def _resize(self, event):
        if event.widget is self.root:
            width = max(400, event.width - 40)
            for label in (self.hint_label, self.time_label, self.note_label, self.status_label):
                label.configure(wraplength=width)
            self.summary_label.configure(wraplength=max(250, event.width - 400))

    def _show_note(self, event=None):
        selected = self.tree.selection()
        if selected:
            values = self.tree.item(selected[0], "values")
            self.selection_note.set(f"{values[0]} — {values[1]}\n{values[4]}")
        else:
            self.selection_note.set("Select a row to see its full review note.")

    def _populate(self, profile):
        self._loading = True
        self._base = profile
        form = new_form() if profile is None else ProfileForm(profile.profile_id, profile.name, profile.source_url,
                    profile.destination, profile.floating_timezone, profile.window_start.isoformat(), profile.window_end.isoformat())
        self._draft_id, self.destination = form.profile_id, form.destination
        for name, value in (("name", form.name), ("url", form.source_url), ("zone", form.timezone_name),
                            ("start", form.start_text), ("end", form.end_text),
                            ("destination", form.destination.display_name if form.destination else "")):
            self.values[name].set(value)
        self.profile_combo.configure(values=[p.name for p in self.profiles])
        if profile is not None:
            self.profile_combo.current(next(i for i, p in enumerate(self.profiles) if p.profile_id == profile.profile_id))
        else:
            self.profile_combo.set("Unsaved profile")
        self.calendar_hint.set("Calendar type not yet checked; choose a calendar or build a preview." if profile else "No calendar selected.")
        self.hint_label.configure(foreground="")
        self._loading = False
        self._changed()
        self._set_busy(False)

    def _select_profile(self, event=None):
        index = self.profile_combo.current()
        if not self.jobs.busy and index >= 0:
            self._populate(self.profiles[index])

    def _changed(self, *args):
        if self._loading:
            return
        self._generation += 1
        self.last_preview = None
        self.tree.delete(*self.tree.get_children())
        self.summary.set("No preview yet.")
        self.selection_note.set("Select a row to see its full review note.")
        self.status.set("Inputs changed. Build a preview to inspect the proposed changes.")

    def _collect(self):
        form = ProfileForm(self._draft_id, self.values["name"].get(), self.values["url"].get(), self.destination,
                           self.values["zone"].get(), self.values["start"].get(), self.values["end"].get())
        try:
            profile = build_profile(form, base=self._base)
        except MappingChangeRequired:
            if not self.dialogs.askyesno("Change profile mapping?", "This changes the source or destination for this profile. Existing imports will not be moved. A source change clears confirmed export coverage. Continue?", parent=self.root):
                return None
            profile = build_profile(form, base=self._base, accept_mapping_change=True)
        self._base = profile
        return profile

    def save_profile(self):
        if self.jobs.busy or self._profile_load_error:
            return
        try:
            profile = self._collect()
            if profile is None:
                return
            if any(p.profile_id != profile.profile_id and p.name.casefold() == profile.name.casefold() for p in self.profiles):
                raise FormError("Choose a unique profile name.")
            profiles = [profile if p.profile_id == profile.profile_id else p for p in self.profiles]
            if all(p.profile_id != profile.profile_id for p in self.profiles):
                profiles.append(profile)
            save_profiles(profiles, self.profiles_path)
            self.profiles = profiles
            self._populate(profile)
            self.status.set("Profile saved. No appointments were changed.")
        except FormError as error:
            self.dialogs.showerror("Check profile", str(error), parent=self.root)
        except Exception:
            self.dialogs.showerror("Profile not saved", "Cannot save profiles. The previous file was not replaced.", parent=self.root)

    def _set_busy(self, busy):
        for entry in self.editable:
            entry.configure(state="disabled" if busy else "normal")
        self.profile_combo.configure(state="disabled" if busy else "readonly")
        for button in (self.new_button, self.pick_button, self.preview_button):
            button.configure(state="disabled" if busy else "normal")
        self.save_button.configure(state="disabled" if busy or self._profile_load_error else "normal")
        if busy:
            self.progress.start(12)
        else:
            self.progress.stop()

    def choose_calendar(self):
        if self.jobs.busy:
            return
        self.jobs.start("picker", self._generation, self.picker)
        self._set_busy(True)
        self.status.set("Choose a calendar in Outlook's folder dialog. Cancel there to keep the current destination.")

    def build_preview(self):
        if self.jobs.busy:
            return
        try:
            profile = self._collect()
        except FormError as error:
            self.dialogs.showerror("Check inputs", str(error), parent=self.root)
            return
        if profile is None:
            return
        self._changed()
        def read():
            choice = self.inspector(profile.destination)
            preview = self.coordinator.preview(profile)
            return choice, preview
        self.jobs.start("preview", self._generation, read)
        self._set_busy(True)
        self.status.set("Reading the source and Outlook calendar… No appointments will be changed.")

    def _show_choice(self, choice):
        if choice.is_default is True:
            self.calendar_hint.set(f"Account / store: {choice.store_name} — default calendar.")
            self.hint_label.configure(foreground="")
        else:
            self.calendar_hint.set(f"Account / store: {choice.store_name} — additional or unverified calendar. Its contribution to Exchange free/busy is not confirmed; prefer the account's default calendar.")
            self.hint_label.configure(foreground="#8c5900")

    def _poll(self):
        result = self.jobs.poll()
        if result is not None:
            self._set_busy(False)
            if self._closing:
                self._destroy()
                return
            if result.generation != self._generation:
                self.status.set("Inputs changed during the read. Build a new preview.")
            elif result.error_type:
                self.status.set("Read failed. No appointments were changed.")
                self.dialogs.showerror("Cannot complete read", read_error_message(result.kind, result.error_type), parent=self.root)
            elif result.kind == "picker":
                if result.value is None:
                    self.status.set("Calendar selection cancelled.")
                else:
                    self.destination = result.value.destination
                    self.values["destination"].set(self.destination.display_name)
                    self._show_choice(result.value)
            else:
                choice, preview = result.value
                self._loading = True
                self.destination = choice.destination
                self.values["destination"].set(choice.destination.display_name)
                self._loading = False
                self.last_preview = preview
                self._show_choice(choice)
                rows, zone = preview_rows(preview)
                for row in rows:
                    self.tree.insert("", "end", values=row)
                counts = {kind: sum(row[0] == kind for row in rows) for kind in ("New", "Update", "Unchanged", "Conflict", "Deletion candidate")}
                self.summary.set(" · ".join(f"{number} {kind.lower()}" for kind, number in counts.items()))
                self.time_hint.set(f"Times shown in {zone}. All-day end dates are exclusive.")
                self.status.set("Preview ready. No appointments were changed. Review notes also explain pending removals.")
        if self._alive:
            self._poll_id = self.root.after(80, self._poll)

    def close(self):
        if self.jobs.busy:
            self._closing = True
            self.status.set("Closing when the current read finishes. If Outlook's folder dialog is open, cancel it there.")
        else:
            self._destroy()

    def _destroy(self):
        self._alive = False
        if getattr(self, "_poll_id", None):
            self.root.after_cancel(self._poll_id)
        self.root.destroy()
