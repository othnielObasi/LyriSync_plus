import sqlite3
from pathlib import Path
from typing import Any, Dict, List, Optional


class PreachInfoDB:
    """Small SQLite helper for sermon/preach metadata."""

    def __init__(self, db_path: str = "lyrisync_preach.db"):
        self.db_path = Path(db_path)
        if self.db_path.parent and str(self.db_path.parent) not in {".", ""}:
            self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._init_db()

    def _conn(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    def _init_db(self) -> None:
        with self._conn() as conn:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS preach_info (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL DEFAULT '',
                    title TEXT NOT NULL DEFAULT '',
                    scriptures TEXT NOT NULL DEFAULT '',
                    inspirations TEXT NOT NULL DEFAULT '',
                    subjects TEXT NOT NULL DEFAULT '',
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                )
                """
            )

    def list_entries(self) -> List[Dict[str, Any]]:
        with self._conn() as conn:
            rows = conn.execute(
                """
                SELECT id, name, title, scriptures, inspirations, subjects, created_at, updated_at
                FROM preach_info
                ORDER BY id DESC
                """
            ).fetchall()
            return [dict(r) for r in rows]

    def create_entry(self, payload: Dict[str, Any]) -> int:
        with self._conn() as conn:
            cur = conn.execute(
                """
                INSERT INTO preach_info(name, title, scriptures, inspirations, subjects, updated_at)
                VALUES(?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                """,
                (
                    (payload.get("name") or "").strip(),
                    (payload.get("title") or "").strip(),
                    (payload.get("scriptures") or "").strip(),
                    (payload.get("inspirations") or "").strip(),
                    (payload.get("subjects") or "").strip(),
                ),
            )
            return int(cur.lastrowid)

    def update_entry(self, row_id: int, payload: Dict[str, Any]) -> None:
        with self._conn() as conn:
            conn.execute(
                """
                UPDATE preach_info
                SET name=?, title=?, scriptures=?, inspirations=?, subjects=?, updated_at=CURRENT_TIMESTAMP
                WHERE id=?
                """,
                (
                    (payload.get("name") or "").strip(),
                    (payload.get("title") or "").strip(),
                    (payload.get("scriptures") or "").strip(),
                    (payload.get("inspirations") or "").strip(),
                    (payload.get("subjects") or "").strip(),
                    int(row_id),
                ),
            )

    def delete_entry(self, row_id: int) -> None:
        with self._conn() as conn:
            conn.execute("DELETE FROM preach_info WHERE id=?", (int(row_id),))

    def get_entry(self, row_id: int) -> Optional[Dict[str, Any]]:
        with self._conn() as conn:
            row = conn.execute(
                """
                SELECT id, name, title, scriptures, inspirations, subjects, created_at, updated_at
                FROM preach_info
                WHERE id=?
                """,
                (int(row_id),),
            ).fetchone()
            return dict(row) if row else None
