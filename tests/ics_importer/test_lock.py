"""Actual OS locking across independent interpreter processes, no Outlook."""

from dataclasses import replace
from pathlib import Path
import subprocess
import sys
from tempfile import TemporaryDirectory
from unittest import TestCase

from ics_lock import RunBusyError, target_lock
from test_profiles import profile


CHILD = '''
import os, sys
sys.path.insert(0, sys.argv[1])
from test_profiles import profile
from ics_lock import target_lock, RunBusyError
try:
    with target_lock(profile(), directory=sys.argv[2]):
        if sys.argv[3] == "exit":
            print("locked", flush=True)
            os._exit(0)
        sys.exit(2)
except RunBusyError:
    sys.exit(0 if sys.argv[3] == "busy" else 3)
'''


class LockTests(TestCase):
    def test_nonblocking_cross_process_exclusion_and_abrupt_exit_release(self):
        test_dir = Path(__file__).resolve().parent
        with TemporaryDirectory(dir=test_dir) as directory:
            with target_lock(profile(), directory=directory):
                child = subprocess.run([sys.executable, "-c", CHILD, str(test_dir), directory, "busy"],
                                       cwd=test_dir.parents[1], capture_output=True, timeout=15)
                self.assertEqual(child.returncode, 0, child.stderr.decode())
                with self.assertRaises(RunBusyError):
                    with target_lock(profile(), directory=directory):
                        self.fail("Same target must be locked")
                different = replace(profile(), destination=replace(profile().destination, folder_id="different"))
                with target_lock(different, directory=directory):
                    pass
            child = subprocess.run([sys.executable, "-c", CHILD, str(test_dir), directory, "exit"],
                                   cwd=test_dir.parents[1], capture_output=True, timeout=15)
            self.assertEqual(child.returncode, 0, child.stderr.decode())
            self.assertEqual(child.stdout.strip(), b"locked")
            with target_lock(profile(), directory=directory):
                pass
            self.assertTrue(list(Path(directory).glob("*.lock")))
