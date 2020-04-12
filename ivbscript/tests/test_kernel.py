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
        pytest.param(None, True, False, marks=pytest.mark.xfail),
        pytest.param('', True, True, marks=pytest.mark.xfail)
    ])
    def test_is_complete(self, code: str, is_complete: bool, indent: bool):
        print(self.kernel)
        print(self.kernel.history_manager)
        status = 'complete' if is_complete else 'incomplete'
        assert not (indent and is_complete), 'cant be both completed and indent'
        expected = {'status': status}
        if indent:
            expected.update({'indent': self.kernel.incomplete_indent})
        results = self.kernel.do_is_complete(code)
        assert results == expected, f'code: {code}, expected: {expected}, results: {results}'
