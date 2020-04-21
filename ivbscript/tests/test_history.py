from unittest import mock

import pytest

from ..history import (DBAlreadyConnected, DBNotConnected,
                       FailedGenerateSessionId, HistoryManager)


class TestHistory:
    history = None

    def setup_method(self, method):
        in_memory_path = ":memory:"
        self.history = HistoryManager(in_memory_path)

    def teardown_method(self, method):
        try:
            self.history.disconnect()
        # Ignore if history already disconnected
        except DBNotConnected:
            pass

    def test_empty_history(self):
        self.history.connect()
        assert len(self.history.tail(10)) == 0, "DB Should be empty"

    def test_reconnection(self):
        self.history.connect()
        with pytest.raises(DBAlreadyConnected):
            self.history.connect()

    def test_history_io(self):
        self.history.connect()
        history_entities_num = 10
        history_entities = [(i, f'some code: #{i}') for i in range(history_entities_num)]
        for history_entity in history_entities:
            self.history.append(*history_entity)
        assert len(self.history.tail(history_entities_num * 10)) == history_entities_num, \
            f"History Should contain {history_entities_num} rows"
        expected_tail = [(self.history.session_id, *history_entity) for history_entity in history_entities]
        tail = self.history.tail(history_entities_num)
        assert tail == expected_tail, 'expected_tail != tail'

    def test_multiple_disconnections(self):
        self.history.connect()
        self.history.disconnect()
        with pytest.raises(DBNotConnected):
            self.history.disconnect()

    def test_session_exists(self):
        with mock.patch.object(HistoryManager, 'is_session_exists', return_value=True) as is_session_exists_mock:
            with pytest.raises(FailedGenerateSessionId):
                self.history.connect()
            assert is_session_exists_mock.call_count != self.history.MAX_SESSION_ID_GENERATE_TRIES + 1
