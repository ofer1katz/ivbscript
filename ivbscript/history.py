import sqlite3
import uuid
from contextlib import closing


class FailedGenerateSessionId(Exception):
    pass


class DBAlreadyConnected(Exception):
    pass


class DBNotConnected(Exception):
    pass


class HistoryManager:
    """
    SQLite DB manager capable of retrieving and appending
    history to a database on disk
    """
    MAX_SESSION_ID_GENERATE_TRIES = 15

    def __init__(self, history_path: str):
        """
        :param history_path: Path to database (created if does not exist)
        :type history_path: string
        """
        self.history_db_path = history_path
        self.history_db = None
        self.session_id = self._generate_session_id()

    @staticmethod
    def _generate_session_id():
        return str(uuid.uuid4())

    @property
    def connected(self):
        try:
            self.history_db.execute('SELECT 1')
            return True
        except (sqlite3.ProgrammingError, AttributeError):
            return False

    def connect(self):
        if self.connected:
            raise DBAlreadyConnected
        self.history_db = sqlite3.connect(self.history_db_path)
        with closing(self.history_db.cursor()) as cursor:
            cursor.execute("""CREATE TABLE IF NOT EXISTS history (
                                          id INTEGER PRIMARY KEY AUTOINCREMENT,
                                          session_id VARCHAR(36) NOT NULL,
                                          line INTEGER NOT NULL,
                                          source TEXT
                                      );
                                   """)
            self.history_db.commit()

        tries_left = self.MAX_SESSION_ID_GENERATE_TRIES
        while self.is_session_exists() and tries_left > 0:
            tries_left -= 1
            self.session_id = self._generate_session_id()

        if self.is_session_exists():
            raise FailedGenerateSessionId

    def is_session_exists(self):
        with closing(self.history_db.cursor()) as cursor:
            result = cursor.execute("""SELECT session_id
            FROM history
            where session_id = ?
            LIMIT 1""", (self.session_id,)).fetchone()
        return bool(result)

    def append(self, line: int, source: str):
        self.history_db.execute("INSERT INTO history (session_id, line, source)"
                                " VALUES (?,?,?)", (self.session_id, line, source))
        self.history_db.commit()

    def tail(self, lines_back: int):
        with closing(self.history_db.cursor()) as cursor:
            return cursor.execute("""
            SELECT h.session_id, h.line, h.source
            FROM history h
            JOIN (
                    SELECT MIN(id) as order_, session_id
                    FROM history
                    GROUP BY session_id
            ) o
            ON o.session_id = h.session_id
            ORDER BY o.order_
            LIMIT ? """, (lines_back,)).fetchall()

    def disconnect(self):
        try:
            self.history_db.close()
        except AttributeError:
            raise DBNotConnected
        finally:
            self.history_db = None
