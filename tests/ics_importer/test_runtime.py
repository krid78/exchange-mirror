"""Runtime diagnostics must stay offline and isolated from user profiles."""
from functools import partial
import json
from pathlib import Path
import sys
from tempfile import TemporaryDirectory
from unittest import TestCase, skipUnless
from unittest.mock import patch

from ics_profiles import default_profiles_path
from ics_runtime import check_runtime


@skipUnless(sys.platform == "win32", "Windows runtime dependencies required")
class RuntimeTests(TestCase):
    def test_real_runtime_check_uses_temporary_profiles_and_never_starts_outlook_or_sync(self):
        with TemporaryDirectory(dir=Path(__file__).parent) as directory:
            directory = Path(directory)
            with patch.dict("os.environ", {"APPDATA": str(directory / "appdata")}):
                existing = default_profiles_path()
                existing.parent.mkdir(parents=True)
                existing.write_bytes(b"existing profile must remain unchanged")
                with (patch("ics_runtime.TemporaryDirectory", partial(TemporaryDirectory, dir=directory)),
                      patch("win32com.client.GetActiveObject", side_effect=AssertionError("No Outlook")) as attach,
                      patch("ics_run.ImportCoordinator.preview", side_effect=AssertionError("No preview")) as preview,
                      patch("ics_run.ImportCoordinator.execute", side_effect=AssertionError("No sync")) as execute):
                    report = directory / "report.json"
                    self.assertEqual(check_runtime(report), 0)
                data = json.loads(report.read_text(encoding="utf-8"))
                self.assertTrue(data["ok"])
                self.assertFalse(data["frozen"])
                self.assertEqual(set(data["checks"]), {"com_imports", "timezone_parser", "profile_roundtrip", "gui_idle"})
                self.assertEqual(existing.read_bytes(), b"existing profile must remain unchanged")
                self.assertFalse(list(directory.glob("ics-importer-runtime-*")))
                attach.assert_not_called()
                preview.assert_not_called()
                execute.assert_not_called()

    def test_missing_timezone_fails_with_static_diagnostic(self):
        with TemporaryDirectory(dir=Path(__file__).parent) as directory:
            report = Path(directory) / "failure.json"
            with patch("ics_runtime.ZoneInfo", side_effect=RuntimeError("PRIVATE PATH")):
                self.assertEqual(check_runtime(report), 1)
            data = json.loads(report.read_text(encoding="utf-8"))
            self.assertFalse(data["ok"])
            self.assertEqual(data["failure"], {"stage": "timezone_parser", "type": "RuntimeError"})
            self.assertNotIn("PRIVATE", report.read_text(encoding="utf-8"))
