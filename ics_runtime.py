"""Opt-in, offline runtime diagnostics for the standalone importer distribution."""

import ctypes
from datetime import date, datetime, timedelta
import json
from pathlib import Path
import sys
from tempfile import TemporaryDirectory
import tkinter as tk
from uuid import UUID
from zoneinfo import ZoneInfo


def check_runtime(report_path):
    """Exercise bundled runtime without user profiles, network or Outlook access.

    Write only a diagnostic report and temporary synthetic profile. The window is
    withdrawn; no preview or synchronization is requested.
    """
    report = {"ok": False, "frozen": bool(getattr(sys, "frozen", False))}
    root = None
    window = None
    stage = "dependencies"
    try:
        import icalendar
        import dateutil
        import tzdata
        import pythoncom
        import pywintypes
        import win32com.client
        import win32timezone
        from ics_gui import ImporterWindow
        from ics_profiles import Destination, ImportProfile, load_profiles, save_profiles
        from ics_source import parse_ics

        report["versions"] = {"icalendar": icalendar.__version__, "dateutil": dateutil.__version__,
                              "tzdata": tzdata.__version__}
        report["elevated"] = bool(ctypes.windll.shell32.IsUserAnAdmin())
        stage = "timezone_parser"
        zone = ZoneInfo("Europe/Berlin")
        if (datetime(2030, 1, 15, tzinfo=zone).utcoffset() != timedelta(hours=1)
                or datetime(2030, 7, 15, tzinfo=zone).utcoffset() != timedelta(hours=2)):
            raise ValueError("Unexpected timezone rules")
        data = (b"BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//Runtime check//EN\r\n"
                b"BEGIN:VEVENT\r\nUID:runtime@example.invalid\r\nDTSTAMP:20300101T000000Z\r\n"
                b"DTSTART;TZID=Europe/Berlin:20300115T140000\r\nDURATION:PT1H\r\n"
                b"SUMMARY:Synthetic runtime check\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n")
        events = parse_ics(data, UUID(int=1))
        if len(events) != 1 or events[0].fields.start.hour != 13:
            raise ValueError("Unexpected parsed event")
        stage = "profiles_gui"
        with TemporaryDirectory(prefix="ics-importer-runtime-") as temporary:
            profiles_path = Path(temporary) / "profiles.json"
            profile = ImportProfile(UUID(int=1), "Runtime check", "https://example.invalid/runtime",
                                    Destination("synthetic-store", "synthetic-folder", "Synthetic"),
                                    "Europe/Berlin", date(2030, 1, 1), date(2030, 2, 1))
            save_profiles((profile,), profiles_path)
            if load_profiles(profiles_path) != (profile,):
                raise ValueError("Profile roundtrip failed")
            root = tk.Tk()
            root.withdraw()
            window = ImporterWindow(root, profiles_path=profiles_path)
            root.update()
            if window.jobs.busy or window.last_preview is not None or window.last_result is not None:
                raise ValueError("Unexpected startup activity")
            if str(window.execute_button.cget("state")) != "disabled":
                raise ValueError("Unexpected enabled synchronization")
            report["checks"] = ["com_imports", "timezone_parser", "profile_roundtrip", "gui_idle"]
            window.close()
            window = root = None
        report["ok"] = True
    except Exception as error:
        report["failure"] = {"stage": stage, "type": type(error).__name__}
    finally:
        if window is not None:
            window.close()
        elif root is not None:
            root.destroy()
    Path(report_path).write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
    return 0 if report["ok"] else 1
