"""Manual preview-only importer entry point; nothing starts when imported."""

import argparse
from pathlib import Path
import tkinter as tk

from ics_gui import ImporterWindow


def main():
    parser = argparse.ArgumentParser(description="Preview ICS calendar synchronization without changing appointments.")
    parser.add_argument("--profiles", type=Path, help="Optional alternate profiles JSON file")
    args = parser.parse_args()
    root = tk.Tk()
    ImporterWindow(root, profiles_path=args.profiles)
    root.mainloop()


if __name__ == "__main__":
    main()
