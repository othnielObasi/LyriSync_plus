import tempfile
import unittest
from pathlib import Path

from lyrisync_plus.preach_info_db import PreachInfoDB


class PreachInfoDBTests(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.TemporaryDirectory()
        self.db_path = Path(self.tmpdir.name) / "data" / "preach.sqlite3"
        self.db = PreachInfoDB(str(self.db_path))

    def tearDown(self):
        self.tmpdir.cleanup()

    def test_create_and_list(self):
        row_id = self.db.create_entry(
            {
                "name": "Pastor Kim",
                "title": "Hope in Trials",
                "scriptures": "Romans 5:1-5",
                "inspirations": "Stay faithful.",
                "subjects": "Hope, Perseverance",
            }
        )
        self.assertGreater(row_id, 0)

        rows = self.db.list_entries()
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["name"], "Pastor Kim")
        self.assertEqual(rows[0]["title"], "Hope in Trials")

    def test_update_and_get(self):
        row_id = self.db.create_entry({"name": "Pastor A", "title": "Old"})
        self.db.update_entry(row_id, {"name": "Pastor B", "title": "New"})

        row = self.db.get_entry(row_id)
        self.assertIsNotNone(row)
        self.assertEqual(row["name"], "Pastor B")
        self.assertEqual(row["title"], "New")

    def test_delete(self):
        row_id = self.db.create_entry({"name": "Pastor Z"})
        self.db.delete_entry(row_id)
        self.assertIsNone(self.db.get_entry(row_id))
        self.assertEqual(self.db.list_entries(), [])


if __name__ == "__main__":
    unittest.main()
