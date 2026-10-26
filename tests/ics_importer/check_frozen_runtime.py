"""Opt-in check of the real portable build with Python/Tcl paths removed."""
import argparse
import json
import os
from pathlib import Path
import shutil
import subprocess
from tempfile import TemporaryDirectory


ROOT = Path(__file__).resolve().parents[2]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--distribution", type=Path, default=ROOT / "dist/ics-importer")
    args = parser.parse_args()
    base = ROOT / ".venv/tmp"
    base.mkdir(parents=True, exist_ok=True)
    with TemporaryDirectory(prefix="frozen-check-", dir=base) as directory:
        directory = Path(directory)
        portable = directory / "Portable importer with spaces"
        shutil.copytree(args.distribution, portable)
        executable = portable / "ics-importer.exe"
        if not executable.is_file():
            raise RuntimeError("Build the importer first")
        env = dict(os.environ)
        for name in ("PYTHONHOME", "PYTHONPATH", "TCL_LIBRARY", "TK_LIBRARY", "VIRTUAL_ENV"):
            env.pop(name, None)
        windows = Path(os.environ["SystemRoot"])
        env["PATH"] = os.pathsep.join((str(windows / "System32"), str(windows)))
        env["PYTHONTZPATH"] = ""  # Exercise bundled tzdata, not a system database.
        env["APPDATA"] = str(directory / "appdata")
        env["TEMP"] = env["TMP"] = str(directory)
        report = directory / "runtime.json"
        result = subprocess.run([str(executable), "--check-runtime", str(report)],
                                cwd=directory, env=env, timeout=45)
        if not report.is_file():
            raise RuntimeError("Frozen runtime produced no diagnostic report")
        data = json.loads(report.read_text(encoding="utf-8"))
        print(json.dumps(data, indent=2))
        if result.returncode != 0 or not data.get("ok") or not data.get("frozen") or data.get("elevated") is not False:
            raise RuntimeError("Frozen runtime check failed or ran elevated")
        if (directory / "appdata").exists():
            raise RuntimeError("Runtime check unexpectedly used user profile storage")
        print("PASS: portable runtime, normal user, offline parser/timezones/profiles, idle GUI")


if __name__ == "__main__":
    main()
