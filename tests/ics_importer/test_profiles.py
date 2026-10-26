"""JSON profile storage in temporary directories only."""

from dataclasses import replace
from datetime import date
import json
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase
from unittest.mock import patch
from uuid import UUID

from ics_profiles import (Destination, ImportProfile, ProfileError, default_profiles_path,
                          load_profiles, save_profiles, update_profile)


def profile():
    return ImportProfile(UUID(int=1), "Synthetic", "https://calendar.example.invalid/feed?key=old",
                         Destination("store", "folder", "Calendar"), "Europe/Berlin",
                         date(2026, 1, 1), date(2027, 1, 1))


class ProfileTests(TestCase):
    def setUp(self):
        self.directory = TemporaryDirectory(dir=Path(__file__).resolve().parent)
        self.addCleanup(self.directory.cleanup)
        self.path = Path(self.directory.name) / "nested" / "profiles.json"

    def test_roundtrip_and_missing_file(self):
        self.assertEqual(load_profiles(self.path), ())
        save_profiles([profile()], self.path)
        self.assertEqual(load_profiles(self.path), (profile(),))
        document = json.loads(self.path.read_text())
        self.assertEqual(document["version"], 1)
        self.assertEqual(document["profiles"][0]["destination"]["folder_id"], "folder")

    def test_url_and_target_changes_require_confirmation_and_retain_uuid(self):
        for options in ({"source_url": profile().source_url + "new"},
                        {"destination": Destination("store", "new", "Other")}):
            with self.subTest(options=options), self.assertRaises(ProfileError):
                update_profile(profile(), **options)
        changed = update_profile(profile(), source_url=profile().source_url + "new", confirm_source_change=True)
        self.assertEqual(changed.profile_id, profile().profile_id)
        changed = update_profile(profile(), destination=Destination("store", "new", "Other"),
                                 confirm_destination_change=True)
        self.assertEqual(changed.profile_id, profile().profile_id)

    def test_invalid_or_duplicate_data_does_not_change_file(self):
        save_profiles([profile()], self.path)
        old = self.path.read_bytes()
        with self.assertRaises(ProfileError):
            save_profiles([profile(), profile()], self.path)
        self.assertEqual(self.path.read_bytes(), old)
        for data in (b"{broken", b'{"version":1,"version":1,"profiles":[]}',
                     b'{"version":true,"profiles":[]}', b'{"version":2,"profiles":[]}'):
            self.path.write_bytes(data)
            with self.subTest(data=data), self.assertRaises(ProfileError):
                load_profiles(self.path)
            self.assertEqual(self.path.read_bytes(), data)

    def test_failed_replace_preserves_previous_file_and_cleans_temp(self):
        save_profiles([profile()], self.path)
        old = self.path.read_bytes()
        with patch("ics_profiles.os.replace", side_effect=OSError("simulated failure")):
            with self.assertRaises(ProfileError):
                save_profiles([replace(profile(), name="New")], self.path)
        self.assertEqual(self.path.read_bytes(), old)
        self.assertEqual(list(self.path.parent.iterdir()), [self.path])

    def test_validation_and_schema(self):
        for options in ({"window_end": profile().window_start}, {"profile_id": "not-a-uuid"},
                        {"floating_timezone": ""}, {"destination": "Calendar"}):
            with self.subTest(options=options), self.assertRaises(ValueError):
                replace(profile(), **options)
        save_profiles([profile()], self.path)
        document = json.loads(self.path.read_text())
        document["profiles"][0]["window_start"] = "20260101"
        self.path.write_text(json.dumps(document))
        with self.assertRaises(ProfileError):
            load_profiles(self.path)

    def test_default_path_without_accessing_real_user_data(self):
        with patch.dict("ics_profiles.os.environ", {"APPDATA": str(self.path.parent)}):
            self.assertEqual(default_profiles_path(), self.path.parent / "Exchange Mirror" / "ics-importer" / "profiles.json")
        with patch.dict("ics_profiles.os.environ", {}, clear=True):
            with self.assertRaises(ProfileError):
                default_profiles_path()
