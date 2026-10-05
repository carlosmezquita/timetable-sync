from __future__ import annotations

import json
import os
import sqlite3
from contextlib import contextmanager
from pathlib import Path

from .models import SyncError


class Store:
    def __init__(self, directory: Path):
        directory.mkdir(parents=True, exist_ok=True)
        if os.name != "nt":
            directory.chmod(0o700)
        self.directory = directory
        self.db = sqlite3.connect(directory / "state.sqlite3")
        self.db.row_factory = sqlite3.Row
        self.db.executescript("""
        PRAGMA journal_mode=WAL;
        CREATE TABLE IF NOT EXISTS metadata (key TEXT PRIMARY KEY, value TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS events (
            key TEXT PRIMARY KEY, data TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS rules (
            selector TEXT PRIMARY KEY, value TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS operations (
            key TEXT PRIMARY KEY, transaction_id TEXT NOT NULL
        );
        """)
        self.db.commit()

    def meta(self, key, default=None):
        row = self.db.execute("SELECT value FROM metadata WHERE key=?", (key,)).fetchone()
        return json.loads(row[0]) if row else default

    def set_meta(self, key, value):
        self.db.execute("INSERT OR REPLACE INTO metadata VALUES (?,?)", (key, json.dumps(value)))
        self.db.commit()

    def bind(self, binding):
        previous = self.meta("binding")
        if previous is not None and previous != binding:
            raise SyncError(
                "This state directory belongs to another source, calendar or account. Use a separate state directory."
            )
        self.set_meta("binding", binding)

    def events(self):
        return {
            row["key"]: json.loads(row["data"]) for row in self.db.execute("SELECT * FROM events")
        }

    def save(self, key, data):
        self.db.execute("INSERT OR REPLACE INTO events VALUES (?,?)", (key, json.dumps(data)))
        self.db.commit()

    def delete(self, key):
        self.db.execute("DELETE FROM events WHERE key=?", (key,))
        self.db.execute("DELETE FROM operations WHERE key=?", (key,))
        self.db.commit()

    def rules(self):
        return dict(self.db.execute("SELECT selector,value FROM rules"))

    def rule(self, selector, value):
        if value == "default":
            self.db.execute("DELETE FROM rules WHERE selector=?", (selector,))
        else:
            self.db.execute("INSERT OR REPLACE INTO rules VALUES (?,?)", (selector, value))
        self.db.commit()

    def transaction(self, key):
        import uuid

        row = self.db.execute(
            "SELECT transaction_id FROM operations WHERE key=?", (key,)
        ).fetchone()
        if row:
            return row[0]
        identifier = str(uuid.uuid4())
        self.db.execute("INSERT INTO operations VALUES (?,?)", (key, identifier))
        self.db.commit()
        return identifier

    @contextmanager
    def lock(self):
        path = self.directory / "sync.lock"
        stream = os.fdopen(os.open(path, os.O_RDWR | os.O_CREAT, 0o600), "r+b", buffering=0)
        try:
            if os.fstat(stream.fileno()).st_size == 0:
                stream.write(b"0")
            stream.seek(0)
            if os.name == "nt":
                import msvcrt

                stream.seek(0)
                msvcrt.locking(stream.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl

                fcntl.flock(stream.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError as exc:
            stream.close()
            raise SyncError("Another synchroniser process is running.") from exc
        try:
            yield
        finally:
            stream.close()
