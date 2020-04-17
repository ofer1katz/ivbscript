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

    def __init__(self, history_path):
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
            self.history_db.execute('select 1')
            return True
        except (sqlite3.ProgrammingError, AttributeError):
            return False

    def connect(self):
        if self.connected:
            raise DBAlreadyConnected
        self.history_db = sqlite3.connect(self.history_db_path)
        with closing(self.history_db.cursor()) as cursor:
            cursor.execute("""CREATE TABLE IF NOT EXISTS history (
                                          session_id varchar(36),
                                          line integer,
                                          source text,
                                          PRIMARY KEY(session_id, line)
                                      );
                                   """)
            self.history_db.commit()

        max_tries = 15
        while self.is_session_exists() and max_tries > 0:
            max_tries -= 1
            self.session_id = uuid.uuid4()

        if self.is_session_exists():
            raise FailedGenerateSessionId

    def is_session_exists(self):
        with closing(self.history_db.cursor()) as cursor:
            result = cursor.execute("SELECT session_id FROM history where session_id = ? LIMIT 1", (self.session_id,)).fetchone()
        return bool(result)

    def append(self, line, source):
        self.history_db.execute("INSERT INTO history VALUES (?,?,?)", (self.session_id, line, source))
        self.history_db.commit()

    def tail(self, lines_back):
        with closing(self.history_db.cursor()) as cursor:
            return cursor.execute("""SELECT session_id, line, source FROM history
                                           ORDER BY session_id, line
                                           LIMIT ? """, (lines_back,)).fetchall()

    def disconnect(self):
        try:
            self.history_db.close()
        except AttributeError:
            raise DBNotConnected
        finally:
            self.history_db = None
