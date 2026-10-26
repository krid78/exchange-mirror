"""Manual ICS importer entry point; nothing starts when imported."""

import argparse
from pathlib import Path
import tkinter as tk

from ics_gui import ImporterWindow


def main():
    parser = argparse.ArgumentParser(description="Review and explicitly synchronize a public ICS feed into an Outlook calendar.")
    parser.add_argument("--profiles", type=Path, help="Optional alternate profiles JSON file")
    parser.add_argument("--check-runtime", type=Path, metavar="REPORT.json",
                        help="Write offline runtime diagnostics; no network or Outlook access")
    args = parser.parse_args()
    if args.check_runtime is not None:
        from ics_runtime import check_runtime
        raise SystemExit(check_runtime(args.check_runtime))
    root = tk.Tk()
    ImporterWindow(root, profiles_path=args.profiles)
    root.mainloop()


if __name__ == "__main__":
    main()
