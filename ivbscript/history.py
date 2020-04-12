import sqlite3
from contextlib import closing


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
        self.session = None

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
                                          session integer,
                                          line integer,
                                          source text,
                                          PRIMARY KEY(session, line)
                                      );
                                   """)
            self.history_db.commit()

            result = cursor.execute("SELECT max(session) FROM history").fetchone()

        self.session = result[0] + 1 if result[0] else 1

    def append(self, line, source):
        self.history_db.execute("INSERT INTO history VALUES (?,?,?)", (self.session, line, source))
        self.history_db.commit()

    def tail(self, lines_back):
        with closing(self.history_db.cursor()) as cursor:
            return cursor.execute("""SELECT session, line, source FROM history
                                           ORDER BY session, line
                                           LIMIT ? """, (lines_back,)).fetchall()

    def disconnect(self):
        try:
            self.history_db.close()
        except AttributeError:
            raise DBNotConnected
        finally:
            self.history_db = None
