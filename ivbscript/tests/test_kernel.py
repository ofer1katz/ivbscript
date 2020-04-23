import os
from typing import Callable, Dict, List
from unittest import mock

import pytest

from ..kernel import VBScriptKernel


class TestKernel:
    kernel = None

    def setup_method(self, method):
        with mock.patch.object(VBScriptKernel, 'run'):
            self.kernel = VBScriptKernel()

    @pytest.mark.parametrize("code,is_complete,indent", [
        ('Sub x()', False, True),
        ('Sub x', False, True),
        ('Sub x()\n\tWScript.Echo 1\nEnd Sub', True, False),
        ('Sub x\n\tWScript.Echo 1\nEnd Sub', True, False),
        ('Sub x WScript.Echo 1 End Sub', True, False),
        ('Sub x WScript.Echo 1 End Sub Function y() y = 2 End Function', True, False),
        ('Function x()', False, True),
        ('Function x()\n\tWScript.Echo 1\nEnd Function', True, False),
        ('Function x()\n\tWScript.Echo 1\n\n', True, False),
        ('Function x() WScript.Echo 1 _', False, False),
        pytest.param(None, True, False, marks=pytest.mark.xfail),
        pytest.param('', True, True, marks=pytest.mark.xfail)
    ])
    def test_is_complete(self, code: str, is_complete: bool, indent: bool):
        status = 'complete' if is_complete else 'incomplete'
        assert not (indent and is_complete), 'cant be both completed and indent'
        expected = {'status': status}
        if indent:
            expected.update({'indent': self.kernel.incomplete_indent})
        results = self.kernel.do_is_complete(code)
        assert results == expected, f'code: {code}, expected: {expected}, results: {results}'

    @pytest.mark.parametrize("code,function", [
        ('exit', VBScriptKernel._terminate_app),
        ('!whoami', VBScriptKernel._handle_command_line_code),
        ('%file', VBScriptKernel._handle_magic),
        ('Dim i: i = 1', VBScriptKernel._handle_vbscript_command),
        ('cls', os.system),
        pytest.param('!cls', os.system, marks=pytest.mark.xfail),
        pytest.param('whoami', VBScriptKernel._handle_command_line_code, marks=pytest.mark.xfail),
        pytest.param('cls', os.remove, marks=pytest.mark.xfail)
    ])
    def test_handle_code_routing(self, code: str, function: Callable[[str], Dict]):

        self.kernel._terminate_app = mock.MagicMock()
        self.kernel._handle_command_line_code = mock.MagicMock()
        with mock.patch.object(self.kernel, '_handle_command_line_code') as _handle_command_line_code_mock:
            with mock.patch.object(self.kernel, '_handle_vbscript_command') as _handle_vbscript_command_mock:
                with mock.patch.object(self.kernel, '_terminate_app') as _terminate_app_mock:
                    with mock.patch.object(self.kernel, '_handle_magic') as _handle_magic_mock:
                        with mock.patch.object(os, 'system') as os_system_mock:
                            mocked_functions = [_terminate_app_mock,
                                                _handle_command_line_code_mock,
                                                _handle_magic_mock,
                                                _handle_vbscript_command_mock,
                                                os_system_mock]
                            self.kernel._handle_code(code)
                            for mocked_function in mocked_functions:
                                if mocked_function._mock_name == function.__name__:
                                    mocked_function.assert_called()
                                    return
        assert False, "Code routing failed. Code wasn't handled"

    @pytest.mark.parametrize("code,cursor_pos,expected_matches,expected_cursor_start", [
        ('En', 2, ['end'], 0),
        ('WScript.Echo vbNe', 17, ['vbNewLine'], 13),
        ('WScript.Echo "x" & vbNe', 23, ['vbNewLine'], 19),
        ('WScript.Echo "x" &vbNe', 22, ['vbNewLine'], 18),
        ('Foo("abc",vbNe', 14, ['vbNewLine'], 10),
        ('Bar(vbNe', 8, ['vbNewLine'], 4),
        ('b', 1, ['ByRef', 'ByVal'], 0),
        ('b', 1, ['ByVal', 'ByRef'], 0),
    ])
    def test_do_complete(self, code: str, cursor_pos: int, expected_matches: List, expected_cursor_start: int):
        """
        :param expected_cursor_start: cursor_pos - completion's initial
        :param expected_matches: a list of variables from pygments.lexers._vbscript_builtins' lists:
                BUILTIN_CONSTANTS, BUILTIN_FUNCTIONS, BUILTIN_VARIABLES, KEYWORDS, OPERATOR_WORDS -
                starts with the completion's initial (ignored case).
        """

        output = self.kernel.do_complete(code, cursor_pos)
        output['matches'].sort()
        expected_matches.sort()
        expected = {
            'matches': expected_matches,
            'cursor_end': cursor_pos,
            'cursor_start': expected_cursor_start,
            'metadata': {},
            'status': 'ok'
        }
        assert output == expected

    def test_do_apply(self):
        with pytest.raises(NotImplementedError):
            self.kernel.do_apply(content=None, bufs=None, msg_id=None, reply_metadata=None)

    def test_do_clear(self):
        with pytest.raises(NotImplementedError):
            self.kernel.do_clear()
