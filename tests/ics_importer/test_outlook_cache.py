"""Change-key snapshots: reuse plain data only when Outlook's token is unchanged."""
from types import SimpleNamespace
from unittest import TestCase
from unittest.mock import Mock

from ics_outlook import OutlookReadError
from test_outlook_read import folder, item, reader, Items, MARKER_KEY
from test_profiles import profile


def table(rows):
    class Table:
        Columns = SimpleNamespace(RemoveAll=Mock(), Add=Mock())
        def __init__(self):
            self.position = 0
        def GetRowCount(self):
            return len(rows)
        @property
        def EndOfTable(self):
            return self.position >= len(rows)
        def GetArray(self, count):
            chunk = rows[self.position:self.position + count]
            self.position += len(chunk)
            return chunk
    return Table()


class ChangeKeyTests(TestCase):
    def test_short_table_ids_are_compared_by_outlook_and_only_long_ids_escape(self):
        adapter = self.setup_adapter()
        self.selected.GetTable.side_effect = lambda: table([("short", b"token")])
        adapter._namespace.GetItemFromID.side_effect = lambda key, store: self.records["one"]
        adapter._namespace.CompareEntryIDs = Mock(return_value=True)
        result = adapter.read_inventory(profile().destination)
        self.assertEqual(result.items[0].entry_id, "one")
        adapter._namespace.CompareEntryIDs.assert_called_once_with("short", "one")
        adapter._namespace.CompareEntryIDs.return_value = False
        self.selected.GetTable.side_effect = lambda: table([("short", b"changed")])
        with self.assertRaises(OutlookReadError):
            adapter.read_inventory(profile().destination)

    def setup_adapter(self):
        self.records = {"one": item(), "manual": item("manual", values={})}
        self.tokens = {key: b"first" for key in self.records}
        self.selected = folder()
        self.selected.GetTable = Mock(side_effect=lambda: table([
            (key, memoryview(self.tokens[key])) for key in self.records]))
        self.adapter = reader(self.selected)
        self.adapter._namespace.GetItemFromID = Mock(side_effect=lambda key, store: self.records[key])
        return self.adapter

    def test_unchanged_records_are_not_reopened_but_changes_additions_and_removals_are_seen(self):
        adapter = self.setup_adapter()
        first = adapter.read_inventory(profile().destination)
        self.assertEqual(adapter._namespace.GetItemFromID.call_count, 2)
        self.assertEqual(adapter.read_inventory(profile().destination), first)
        self.assertEqual(adapter._namespace.GetItemFromID.call_count, 2)
        self.records["one"].Subject = "Changed"
        self.tokens["one"] = b"changed"
        del self.records["manual"]
        self.records["new"] = item("new", values={})
        self.tokens["new"] = b"new"
        current = adapter.read_inventory(profile().destination)
        self.assertEqual(adapter._namespace.GetItemFromID.call_count, 4)
        self.assertEqual(current.items[0].fields.subject, "Changed")
        self.assertEqual({value.entry_id for value in current.items}, {"one", "new"})

    def test_changed_corrupt_markers_cannot_hide_in_cache(self):
        adapter = self.setup_adapter()
        adapter.read_inventory(profile().destination)
        self.records["one"] = item(values={MARKER_KEY: "incomplete"})
        self.tokens["one"] = b"corrupt"
        with self.assertRaises(OutlookReadError) as caught:
            adapter.read_inventory(profile().destination)
        self.assertEqual(caught.exception.code, "inventory_invalid")

    def test_change_during_read_retries_only_discarded_snapshot(self):
        adapter = self.setup_adapter()
        old, new = [("one", b"old")], [("one", b"new")]
        self.selected.GetTable.side_effect = [table(old), table(new), table(new), table(new)]
        result = adapter.read_inventory(profile().destination)
        self.assertEqual(len(result.items), 1)
        self.assertEqual(adapter._namespace.GetItemFromID.call_count, 2)
        self.records["one"].Save.assert_not_called()
        self.records["one"].Delete.assert_not_called()

    def test_continuous_changes_stop_after_three_attempts_without_caching_partial_data(self):
        adapter = self.setup_adapter()
        self.selected.GetTable.side_effect = [table([("one", bytes([index]))]) for index in range(1, 7)]
        with self.assertRaises(OutlookReadError) as caught:
            adapter.read_inventory(profile().destination)
        self.assertEqual(caught.exception.code, "inventory_changed")
        self.assertEqual(self.selected.GetTable.call_count, 6)
        self.assertFalse(adapter._inventory_cache)

    def test_missing_change_keys_fall_back_to_complete_marker_validation(self):
        adapter = self.setup_adapter()
        self.selected.GetTable.side_effect = lambda: table([("one", None)])
        self.selected.Items = Items([item(values={MARKER_KEY: "incomplete"})])
        with self.assertRaises(OutlookReadError) as caught:
            adapter.read_inventory(profile().destination)
        self.assertEqual(caught.exception.code, "inventory_invalid")
        self.assertFalse(adapter._inventory_cache)

    def test_limits_and_duplicate_rows_never_yield_a_partial_inventory(self):
        adapter = self.setup_adapter()
        with self.assertRaises(OutlookReadError) as caught:
            adapter.read_inventory(profile().destination, max_items=1)
        self.assertEqual(caught.exception.code, "inventory_limit")
        self.selected.GetTable.side_effect = lambda: table([("one", b"1"), ("ONE", b"1")])
        with self.assertRaises(OutlookReadError):
            adapter.read_inventory(profile().destination)
        self.assertFalse(adapter._inventory_cache)

    def test_cache_is_discarded_at_session_close(self):
        adapter = self.setup_adapter()
        adapter.read_inventory(profile().destination)
        adapter.close()
        self.assertFalse(adapter._inventory_cache)
        with self.assertRaises(OutlookReadError):
            adapter.read_inventory(profile().destination)
