"""
Exchange Meeting Mirror — System Tray Entry Point

Runs as a persistent tray icon with periodic background sync.
Right-click menu: Start/Stop Sync, Settings, Exit.
"""

import ctypes
import logging
import os
import queue
import threading
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk
import webbrowser
from datetime import datetime, date as date_type
from tkcalendar import DateEntry

import pythoncom
import pystray
from PIL import Image, ImageDraw

import mirror

log = logging.getLogger("mirror")


def create_icon_image():
    """Generate a 64x64 blue circle icon with Pillow."""
    size = 64
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    # Blue circle with white border
    draw.ellipse([2, 2, size - 3, size - 3], fill=(50, 120, 220), outline=(255, 255, 255), width=2)
    # Draw two small arrows (sync symbol) as simple lines
    mid = size // 2
    # Right arrow
    draw.line([mid - 10, mid - 6, mid + 10, mid - 6], fill="white", width=2)
    draw.line([mid + 5, mid - 11, mid + 10, mid - 6, mid + 5, mid - 1], fill="white", width=2)
    # Left arrow
    draw.line([mid + 10, mid + 6, mid - 10, mid + 6], fill="white", width=2)
    draw.line([mid - 5, mid + 1, mid - 10, mid + 6, mid - 5, mid + 11], fill="white", width=2)
    return img


class _TkTextHandler(logging.Handler):
    """Logging handler that appends formatted records into a Text widget.

    Queues the update instead of touching the widget directly, since emit()
    runs on whatever thread is logging (a background sync thread) — calling
    Tk from a non-mainloop thread is the known-unsafe pattern behind the
    tcl86t.dll crashes. The queue is drained only by
    SettingsDialog._pump_ui_queue, which runs on the mainloop thread.
    """

    def __init__(self, ui_queue, text_widget):
        super().__init__()
        self._ui_queue = ui_queue
        self._widget = text_widget
        self.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))

    def emit(self, record):
        msg = self.format(record)

        def _append():
            self._widget.config(state="normal")
            self._widget.insert("end", msg + "\n")
            self._widget.see("end")
            self._widget.config(state="disabled")

        self._ui_queue.put(_append)


class SettingsDialog:
    """Tabbed settings dialog for Exchange Mirror.

    Reads config on open, writes back on Save. Safe to run in a daemon thread.
    """

    def __init__(self, controller):
        self._controller = controller
        self.root = tk.Tk()
        self.root.title("Exchange Mirror — Settings")
        self.root.resizable(False, False)
        self.root.protocol("WM_DELETE_WINDOW", self._on_cancel)

        # Background threads (Detect, Test, Sync Now) must never call Tk
        # methods directly — only this queue, drained by _pump_ui_queue on
        # the mainloop thread.
        self._ui_queue = queue.Queue()
        self.root.after(50, self._pump_ui_queue)

        try:
            cfg = mirror._load_config()
        except Exception as exc:
            self.root.withdraw()
            messagebox.showerror("Error", f"Cannot load config:\n{exc}", parent=self.root)
            self.root.destroy()
            return

        # --- tk variables for all config fields ---
        self._v_interval   = tk.IntVar(value=cfg.SYNC_INTERVAL_MINUTES)
        self._v_past       = tk.IntVar(value=cfg.SYNC_PAST_DAYS)
        self._v_future     = tk.IntVar(value=cfg.SYNC_FUTURE_DAYS)
        self._v_chunk      = tk.IntVar(value=cfg.SYNC_CHUNK_DAYS)
        # "relative" = use past/future days; "fixed" = use date pickers
        self._v_mode       = tk.StringVar(value="fixed" if cfg.SYNC_START_DATE else "relative")
        self._init_start   = cfg.SYNC_START_DATE or ""
        self._init_end     = cfg.SYNC_END_DATE or ""
        self._v_prefix     = tk.StringVar(value=cfg.MIRROR_SUBJECT_PREFIX)
        self._v_category   = tk.StringVar(value=cfg.MIRROR_CATEGORY)
        self._v_dry_run    = tk.BooleanVar(value=cfg.DRY_RUN)
        self._v_startup    = tk.BooleanVar(value=mirror.is_startup_enabled())
        self._v_log_file   = tk.StringVar(value=cfg.LOG_FILE)
        self._v_account_a  = tk.StringVar(value=cfg.ACCOUNT_A or "")
        self._v_account_b  = tk.StringVar(value=cfg.ACCOUNT_B or "")

        self._build_ui()

    def _pump_ui_queue(self):
        """Drain callbacks queued by background threads. Runs only on the
        mainloop thread, so this is the only place that touches Tk after
        startup other than the mainloop itself."""
        while True:
            try:
                fn = self._ui_queue.get_nowait()
            except queue.Empty:
                break
            try:
                fn()
            except Exception:
                log.exception("UI update failed")
        self.root.after(50, self._pump_ui_queue)

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self):
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True, padx=10, pady=(10, 0))

        self._build_general_tab(notebook)
        self._build_accounts_tab(notebook)
        self._build_advanced_tab(notebook)

        # Bottom button bar
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill="x", padx=10, pady=8)
        ttk.Button(btn_frame, text="Save", command=self._on_save, width=10).pack(side="left", padx=(0, 4))
        ttk.Button(btn_frame, text="Cancel", command=self._on_cancel, width=10).pack(side="left")

    def _build_general_tab(self, notebook):
        tab = ttk.Frame(notebook, padding=10)
        notebook.add(tab, text="General")

        vcmd_int = (tab.register(self._validate_int), "%P")

        # Sync Interval (always active)
        interval_f = ttk.LabelFrame(tab, text="Sync Interval", padding=8)
        interval_f.pack(fill="x", pady=(0, 8))
        interval_f.columnconfigure(1, weight=1)
        self._row(interval_f, "Interval (min):", ttk.Entry(interval_f, textvariable=self._v_interval,
                  width=8, validate="key", validatecommand=vcmd_int), 0)

        # Sync Window (radio-switched)
        window_f = ttk.LabelFrame(tab, text="Sync Window", padding=8)
        window_f.pack(fill="x", pady=(0, 8))
        window_f.columnconfigure(1, weight=1)

        # Radio buttons
        today = date_type.today()
        radio_frame = ttk.Frame(window_f)
        radio_frame.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))
        ttk.Radiobutton(radio_frame, text="Fixed date range", variable=self._v_mode,
                        value="fixed",    command=self._toggle_window_mode).pack(side="left", padx=(0, 16))
        ttk.Radiobutton(radio_frame, text="Relative window", variable=self._v_mode,
                        value="relative", command=self._toggle_window_mode).pack(side="left")
        ttk.Label(radio_frame, text=f"(Today: {today:%Y-%m-%d})",
                  foreground="gray").pack(side="left", padx=(12, 0))

        # Relative widgets
        self._lbl_past   = ttk.Label(window_f, text="Past days:")
        self._ent_past   = ttk.Entry(window_f, textvariable=self._v_past, width=8,
                                     validate="key", validatecommand=vcmd_int)
        self._lbl_future = ttk.Label(window_f, text="Future days:")
        self._ent_future = ttk.Entry(window_f, textvariable=self._v_future, width=8,
                                     validate="key", validatecommand=vcmd_int)
        self._lbl_past.grid  (row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        self._ent_past.grid  (row=1, column=1, sticky="w", pady=4)
        self._lbl_future.grid(row=2, column=0, sticky="w", padx=(0, 8), pady=4)
        self._ent_future.grid(row=2, column=1, sticky="w", pady=4)

        # Fixed date widgets

        def _parse_date(s):
            try:
                return datetime.strptime(s, "%Y-%m-%d").date() if s else today
            except ValueError:
                return today

        self._lbl_start = ttk.Label(window_f, text="Start date:")
        self._de_start  = DateEntry(window_f, width=12, date_pattern="yyyy-mm-dd",
                                    year=_parse_date(self._init_start).year,
                                    month=_parse_date(self._init_start).month,
                                    day=_parse_date(self._init_start).day)
        self._lbl_end   = ttk.Label(window_f, text="End date:")
        self._de_end    = DateEntry(window_f, width=12, date_pattern="yyyy-mm-dd",
                                    year=_parse_date(self._init_end).year,
                                    month=_parse_date(self._init_end).month,
                                    day=_parse_date(self._init_end).day)
        self._lbl_start.grid(row=3, column=0, sticky="w", padx=(0, 8), pady=4)
        self._de_start.grid (row=3, column=1, sticky="w", pady=4)
        self._lbl_end.grid  (row=4, column=0, sticky="w", padx=(0, 8), pady=4)
        self._de_end.grid   (row=4, column=1, sticky="w", pady=4)

        # Chunk size (applies to both window modes, so it stays outside the
        # relative/fixed toggle group)
        self._row(window_f, "Chunk size (days):", ttk.Entry(window_f, textvariable=self._v_chunk,
                  width=8, validate="key", validatecommand=vcmd_int), 5)
        ttk.Label(window_f, text="0 = disable chunking", foreground="gray").grid(
            row=6, column=1, sticky="w")

        # Apply initial state
        self._toggle_window_mode()

        # Mirror Markers
        markers = ttk.LabelFrame(tab, text="Mirror Markers", padding=8)
        markers.pack(fill="x", pady=(0, 8))
        markers.columnconfigure(1, weight=1)

        self._row(markers, "Subject prefix:", ttk.Entry(markers, textvariable=self._v_prefix,   width=20), 0)
        self._row(markers, "Category:",       ttk.Entry(markers, textvariable=self._v_category, width=20), 1)

        # Dry Run
        ttk.Checkbutton(tab, text="Dry Run Mode (log changes without applying them)",
                        variable=self._v_dry_run).pack(anchor="w")

        # Startup
        ttk.Checkbutton(tab, text="Start Exchange Mirror when Windows starts",
                        variable=self._v_startup).pack(anchor="w")

    def _build_accounts_tab(self, notebook):
        tab = ttk.Frame(notebook, padding=10)
        notebook.add(tab, text="Accounts")

        ttk.Label(tab, text="Leave empty to auto-detect (requires exactly 2 Outlook accounts).",
                  foreground="gray").pack(anchor="w", pady=(0, 8))

        form = ttk.LabelFrame(tab, text="Account Selection", padding=8)
        form.pack(fill="x", pady=(0, 8))
        form.columnconfigure(1, weight=1)

        self._cb_a = ttk.Combobox(form, textvariable=self._v_account_a, width=36)
        self._cb_b = ttk.Combobox(form, textvariable=self._v_account_b, width=36)
        self._row(form, "Account A (SMTP or name):", self._cb_a, 0)
        self._row(form, "Account B (SMTP or name):", self._cb_b, 1)

        actions = ttk.Frame(tab)
        actions.pack(fill="x", pady=(0, 6))
        ttk.Button(actions, text="Detect from Outlook", command=self._detect_accounts).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Test Connection",     command=self._test_connection).pack(side="left")

        self._lbl_status = ttk.Label(tab, text="")
        self._lbl_status.pack(anchor="w")

    def _build_advanced_tab(self, notebook):
        tab = ttk.Frame(notebook, padding=10)
        notebook.add(tab, text="Advanced")

        log_frame = ttk.LabelFrame(tab, text="Logging", padding=8)
        log_frame.pack(fill="x", pady=(0, 8))
        log_frame.columnconfigure(1, weight=1)

        log_entry = ttk.Entry(log_frame, textvariable=self._v_log_file, width=28)
        self._row(log_frame, "Log file name:", log_entry, 0)
        ttk.Button(log_frame, text="Open Log", command=self._open_log).grid(
            row=0, column=2, padx=(6, 0), pady=4)

        sync_frame = ttk.LabelFrame(tab, text="Manual Sync", padding=8)
        sync_frame.pack(fill="x")
        ttk.Button(sync_frame, text="Sync Now", command=self._sync_now).pack(anchor="w")
        self._lbl_sync = ttk.Label(sync_frame, text="")
        self._lbl_sync.pack(anchor="w", pady=(4, 0))
        self._txt_sync_log = scrolledtext.ScrolledText(
            sync_frame, height=8, state="disabled", wrap="word")
        self._txt_sync_log.pack(fill="both", expand=True, pady=(4, 0))

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _row(parent, label, widget, row):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=4)
        widget.grid(row=row, column=1, sticky="ew", pady=4)

    @staticmethod
    def _validate_int(value):
        return value == "" or value.lstrip("-").isdigit()

    def _set_status(self, text, color="black"):
        self._lbl_status.config(text=text, foreground=color)

    def _toggle_window_mode(self):
        if self._v_mode.get() == "relative":
            for w in (self._lbl_past, self._ent_past, self._lbl_future, self._ent_future):
                w.config(state="normal")
            for w in (self._lbl_start, self._de_start, self._lbl_end, self._de_end):
                w.config(state="disabled")
        else:
            for w in (self._lbl_past, self._ent_past, self._lbl_future, self._ent_future):
                w.config(state="disabled")
            for w in (self._lbl_start, self._de_start, self._lbl_end, self._de_end):
                w.config(state="normal")

    # ------------------------------------------------------------------
    # Button handlers
    # ------------------------------------------------------------------

    def _detect_accounts(self):
        self._set_status("Detecting accounts…", "gray")
        threading.Thread(target=self._do_detect, daemon=True).start()

    def _do_detect(self):
        import pythoncom
        pythoncom.CoInitialize()
        try:
            outlook = mirror.get_outlook()
            accounts = mirror.get_outlook_accounts(outlook)
            values = [f"{smtp}" if not name else f"{name} ({smtp})"
                      for name, smtp, _ in accounts]
            def update():
                self._cb_a["values"] = values
                self._cb_b["values"] = values
                self._set_status(f"Found {len(accounts)} account(s).", "green")
            self._ui_queue.put(update)
        except Exception as exc:
            self._ui_queue.put(lambda: self._set_status(f"Error: {exc}", "red"))
        finally:
            pythoncom.CoUninitialize()

    def _test_connection(self):
        self._set_status("Testing connection…", "gray")
        threading.Thread(target=self._do_test, daemon=True).start()

    def _do_test(self):
        import pythoncom
        pythoncom.CoInitialize()
        try:
            mirror.get_outlook()
            self._ui_queue.put(lambda: self._set_status("Connected \u2713", "green"))
        except Exception as exc:
            self._ui_queue.put(lambda: self._set_status(f"Cannot connect: {exc}", "red"))
        finally:
            pythoncom.CoUninitialize()

    def _open_log(self):
        log_path = os.path.join(mirror.DATA_DIR, self._v_log_file.get() or "mirror.log")
        if not os.path.isfile(log_path):
            messagebox.showwarning("Not Found", f"Log file not found:\n{log_path}", parent=self.root)
            return
        os.startfile(log_path)

    def _sync_now(self):
        self._lbl_sync.config(text="Syncing…", foreground="gray")
        self._txt_sync_log.config(state="normal")
        self._txt_sync_log.delete("1.0", "end")
        self._txt_sync_log.config(state="disabled")

        handler = _TkTextHandler(self._ui_queue, self._txt_sync_log)
        handler.setLevel(logging.INFO)
        logging.getLogger("mirror").addHandler(handler)

        def _run():
            try:
                self._controller.run_once()
                self._ui_queue.put(lambda: self._lbl_sync.config(text="Sync triggered.", foreground="green"))
            finally:
                logging.getLogger("mirror").removeHandler(handler)
        threading.Thread(target=_run, daemon=True).start()

    # ------------------------------------------------------------------
    # Save / Cancel
    # ------------------------------------------------------------------

    def _on_save(self):
        try:
            updates = {
                "SYNC_INTERVAL_MINUTES": self._validated_int("Interval", self._v_interval, min_val=1),
                "SYNC_PAST_DAYS":        self._validated_int("Past days", self._v_past,    min_val=0),
                "SYNC_FUTURE_DAYS":      self._validated_int("Future days", self._v_future, min_val=0),
                "SYNC_CHUNK_DAYS":       self._validated_int("Chunk size", self._v_chunk,  min_val=0),
                "SYNC_START_DATE":       "" if self._v_mode.get() == "relative" else self._de_start.get_date().strftime("%Y-%m-%d"),
                "SYNC_END_DATE":         "" if self._v_mode.get() == "relative" else self._de_end.get_date().strftime("%Y-%m-%d"),
                "MIRROR_SUBJECT_PREFIX": self._v_prefix.get(),
                "MIRROR_CATEGORY":       self._v_category.get(),
                "DRY_RUN":               self._v_dry_run.get(),
                "START_ON_WINDOWS_LOGIN": self._v_startup.get(),
                "LOG_FILE":              self._v_log_file.get().strip() or "mirror.log",
                "ACCOUNT_A":             self._extract_account(self._v_account_a.get()),
                "ACCOUNT_B":             self._extract_account(self._v_account_b.get()),
            }
        except ValueError as exc:
            messagebox.showerror("Validation Error", str(exc), parent=self.root)
            return

        try:
            mirror.save_config(updates)
        except Exception as exc:
            messagebox.showerror("Save Error", str(exc), parent=self.root)
            return

        try:
            mirror.set_startup_enabled(updates["START_ON_WINDOWS_LOGIN"])
        except Exception as exc:
            messagebox.showwarning(
                "Startup Setting",
                f"Settings were saved, but the Windows startup entry could not be updated:\n{exc}",
                parent=self.root,
            )

        messagebox.showinfo("Saved", "Settings saved successfully.", parent=self.root)
        self.root.destroy()

    def _on_cancel(self):
        self.root.destroy()

    @staticmethod
    def _extract_account(val):
        """Strip display name from 'Display Name (smtp@example.com)' → 'smtp@example.com'."""
        val = val.strip()
        if val.endswith(")"):
            idx = val.rfind("(")
            if idx >= 0:
                return val[idx + 1:-1].strip()
        return val

    def _validated_int(self, label, var, min_val=0):
        try:
            val = int(var.get())
        except (tk.TclError, ValueError):
            raise ValueError(f"{label} must be a whole number.")
        if val < min_val:
            raise ValueError(f"{label} must be \u2265 {min_val}.")
        return val

    def run(self):
        """Block until the dialog is closed."""
        self.root.mainloop()


def show_settings(controller):
    """Open the settings dialog (runs in a daemon thread)."""
    try:
        SettingsDialog(controller).run()
    except Exception:
        log.exception("Settings dialog error")


def show_about():
    """Display About dialog with author info and clickable links."""
    text = (
        "Exchange Mirror v1.0\n"
        "\n"
        "Bidirectional Outlook calendar sync.\n"
        "\n"
        "Author: Markus Begerow\n"
"Website: https://www.markus-begerow.de/linktree\n"
    )
    # MB_YESNO (4) + MB_ICONINFORMATION (0x40) = 0x44
    # Yes=6 opens the website
    result = ctypes.windll.user32.MessageBoxW(
        0, text + "\nOpen website?", "About Exchange Mirror", 0x44
    )
    if result == 6:  # IDYES
        webbrowser.open("https://www.markus-begerow.de/linktree")


class SyncController:
    """Manages the background sync thread."""

    def __init__(self):
        self._stop_event = threading.Event()
        self._thread = None
        self._lock = threading.Lock()
        self.running = False

    def start(self):
        with self._lock:
            if self.running:
                return
            self._stop_event.clear()
            self._thread = threading.Thread(target=self._sync_loop, daemon=True)
            self._thread.start()
            self.running = True
            log.info("Sync started.")

    def stop(self):
        with self._lock:
            if not self.running:
                return
            self._stop_event.set()
            self.running = False
            log.info("Sync stopping...")

    def shutdown(self):
        self.stop()
        if self._thread is not None:
            self._thread.join(timeout=10)

    def run_once(self):
        """Trigger a single sync cycle outside the interval loop."""
        def _go():
            pythoncom.CoInitialize()
            try:
                self._run_one_sync()
            finally:
                pythoncom.CoUninitialize()
        threading.Thread(target=_go, daemon=True).start()

    def _sync_loop(self):
        pythoncom.CoInitialize()
        try:
            while not self._stop_event.is_set():
                self._run_one_sync()
                cfg = mirror._load_config()
                interval = getattr(cfg, "SYNC_INTERVAL_MINUTES", 15) * 60
                self._stop_event.wait(timeout=interval)
        finally:
            pythoncom.CoUninitialize()

    def _run_one_sync(self):
        try:
            cfg = mirror._load_config()

            # Fresh COM object each cycle
            outlook = mirror.get_outlook()
            accounts = mirror.get_outlook_accounts(outlook)

            if len(accounts) < 2:
                log.error("Need at least 2 accounts, found %d.", len(accounts))
                return

            account_a_hint = getattr(cfg, "ACCOUNT_A", "")
            account_b_hint = getattr(cfg, "ACCOUNT_B", "")
            acct_a, acct_b = mirror.resolve_accounts(accounts, account_a_hint, account_b_hint)

            name_a, smtp_a, account_a = acct_a
            name_b, smtp_b, account_b = acct_b

            log.info("Syncing between: %s <-> %s", name_a, name_b)
            if cfg.DRY_RUN:
                log.info("*** DRY RUN MODE ***")

            # Reload mirror's config so sync functions use fresh values
            mirror.config = cfg
            mirror.sync_one_direction(account_a, account_b, smtp_a)
            mirror.sync_one_direction(account_b, account_a, smtp_b)

            log.info("Sync complete at %s.", datetime.now().strftime("%H:%M:%S"))

        except Exception:
            log.exception("Sync cycle failed")


_STOP_SETTINGS_WORKER = object()


def _settings_worker(controller, requests):
    """Processes Settings-dialog-open requests, one at a time, on whichever
    thread calls this — which must always be the process's real main
    thread. Tk/Tcl reliably crashes the process at interpreter shutdown
    (with no catchable Python exception) if a Tk interpreter was ever
    created on a non-main thread, even if that thread already exited
    cleanly — so the tray icon loop runs detached instead (see main()),
    leaving the main thread free for this loop.
    """
    while True:
        req = requests.get()
        if req is _STOP_SETTINGS_WORKER:
            return
        show_settings(controller)


def main():
    mirror.init()
    controller = SyncController()

    settings_requests = queue.Queue()

    def on_toggle(icon, item):
        if controller.running:
            controller.stop()
        else:
            controller.start()
        icon.update_menu()

    def on_settings(icon, item):
        settings_requests.put(None)

    def on_about(icon, item):
        threading.Thread(target=show_about, daemon=True).start()

    def on_exit(icon, item):
        controller.shutdown()
        settings_requests.put(_STOP_SETTINGS_WORKER)
        icon.stop()

    def toggle_text(item):
        return "Stop Sync" if controller.running else "Start Sync"

    menu = pystray.Menu(
        pystray.MenuItem(toggle_text, on_toggle, default=True),
        pystray.MenuItem("Settings...", on_settings),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("About...", on_about),
        pystray.MenuItem("Exit", on_exit),
    )

    icon = pystray.Icon(
        name="exchange-mirror",
        icon=create_icon_image(),
        title="Exchange Mirror",
        menu=menu,
    )

    # Run the tray icon loop on its own thread so the main thread stays
    # free for Tk (see _settings_worker docstring for why that's required).
    icon.run_detached()
    _settings_worker(controller, settings_requests)


if __name__ == "__main__":
    main()
