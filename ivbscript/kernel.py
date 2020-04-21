"""
Jupyter kernel implementation for VBScript
"""
import os
import random
import re
import shlex
import time
import traceback
from distutils.spawn import find_executable
from subprocess import PIPE, Popen, TimeoutExpired
from typing import Dict

import psutil
import termcolor
import win32clipboard
from ipykernel.kernelbase import Kernel

from .history import HistoryManager

__version__ = '1.0.0'


class VBScriptKernel(Kernel):
    """

    """
    implementation = 'iVBScript'
    language = "vbscript"
    implementation_version = __version__
    banner = termcolor.colored(r'''
d8b 888     888 888888b.    .d8888b.                   d8b          888
Y8P 888     888 888  "88b  d88P  Y88b                  Y8P          888
    888     888 888  .88P  Y88b.                                    888
888 Y88b   d88P 8888888K.   "Y888b.    .d8888b 888d888 888 88888b.  888888
888  Y88b d88P  888  "Y88b     "Y88b. d88P"    888P"   888 888 "88b 888
888   Y88o88P   888    888       "888 888      888     888 888  888 888
888    Y888P    888   d88P Y88b  d88P Y88b.    888     888 888 d88P Y88b.
888     Y8P     8888888P"   "Y8888P"   "Y8888P 888     888 88888P"   "Y888
                                                           888
                                                           888
                                                           888
    ''', color=random.choice(list(termcolor.COLORS)))
    INTERPRETER = 'cscript.exe'
    COMMAND_LINE_TIMEOUT = 15
    incomplete_indent = '  '
    completion_regexes = {
        'sub': {'start_pattern': r'(^|\s)((private|public)\s+)?sub(\s+)[a-z_][a-z0-9_]*(\(.+\))?',
                'end_pattern': r'(^|\s)end(\s+)sub(\s|$)'},
        'function': {'start_pattern': r'(^|\s)((private|public)\s+)?function(\s+)[a-z_][a-z0-9_]*(\(.+\))?',
                     'end_pattern': r'(^|\s)end(\s+)function(\s|$)'},
        'if': {'start_pattern': r'(^|\s)if(\s+).+(\s+)then(\s|$)', 'end_pattern': r'(^|\s)end(\s+)if(\s|$)'},
        'select': {'start_pattern': r'(^|\s)select(\s+)case', 'end_pattern': r'(^|\s)end(\s+)select(\s|$)'},
        'for': {'start_pattern': r'(^|\s)for(\s+)', 'end_pattern': r'(^|\s)next(\s|$)'},
        'do': {'start_pattern': r'(^|\s)do(\s+)', 'end_pattern': r'(^|\s)loop(\s|$)'},
        'with': {'start_pattern': r'(^|\s)with(\s+)', 'end_pattern': r'(^|\s)end(\s+)with(\s|$)'},
        'property': {'start_pattern': r'(^|\s)((private|public)\s+)?property(\s+)[a-z_][a-z0-9_]*(\(.+\))?',
                     'end_pattern': r'(^|\s)end(\s+)property(\s|$)'},
        'class': {'start_pattern': r'(^|\s)class(\s+)', 'end_pattern': r'(^|\s)end(\s+)class(\s|$)'}
    }

    @property
    def language_info(self):
        return {
            "name": self.language,
            "file_extension": ".vbs",
            "pygments_lexer": "vbscript",
        }

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        assert find_executable(self.INTERPRETER), f'Could not find {self.INTERPRETER}'
        os.chdir(os.path.dirname(os.path.abspath(__file__)))
        runtime_data_dir = os.path.join(os.getcwd(), 'runtime_data')
        if not os.path.exists(runtime_data_dir):
            os.mkdir(runtime_data_dir)
        pid = os.getpid()

        self.history_manager = HistoryManager(self.get_history_path())
        self.cscript = None
        self.stdout_pos = 0
        self.stdout_file_path = os.path.join(runtime_data_dir, f'{pid}.stdout')
        self.stderr_file_path = os.path.join(runtime_data_dir, f'{pid}.stderr')
        self.input_file_path = os.path.join(runtime_data_dir, f'{pid}.input')
        debug_log = os.path.join(runtime_data_dir, f'{pid}.log')

        os.environ.update({'IVBS_CMD_PATH': self.input_file_path, 'IVBS_RET_PATH': self.stderr_file_path,
                           'IVBS_DEBUG_PATH': debug_log})
        self.run()

    @classmethod
    def get_history_path(cls):
        """
        Get platform-specific path to past sessions execution history
        """
        return os.path.join(os.path.expanduser("~"), f".{cls.implementation.lower()}_history.db")

    def run(self):
        self.history_manager.connect()
        with open(self.stdout_file_path, 'w') as stdout_file:
            self.cscript = Popen([
                self.INTERPRETER,
                '//nologo',
                'interpreter.vbs'
            ], stdout=stdout_file, stderr=stdout_file, shell=False, env=os.environ.copy())

    def _get_stdout(self) -> str:
        with open(self.stdout_file_path, 'r') as stdout_file:
            stdout_file.seek(self.stdout_pos)
            data = stdout_file.read()
            self.stdout_pos = stdout_file.tell()
            return data

    def _handle_command_line_code(self, code: str) -> Dict:
        try:
            process = Popen(shlex.split(code), stderr=PIPE, stdout=PIPE)
            try:
                stdout, stderr = process.communicate(timeout=self.COMMAND_LINE_TIMEOUT)
                return {'stdout': stdout.decode('utf-8'), 'stderr': stderr.decode('utf-8')}
            finally:
                process.terminate()
        except (FileNotFoundError, TimeoutExpired) as exception:
            return {'stdout': '', 'stderr': (''.join(traceback.format_exception(None, exception, None)))}

    def _send_command(self, code: str):
        if os.path.exists(self.stderr_file_path):
            os.remove(self.stderr_file_path)
        with open(self.input_file_path, 'w', encoding='utf-8') as input_file:
            input_file.write("\n".join(code.splitlines()))

    def _handle_vbscript_command(self, code: str) -> Dict:
        self._send_command(code)
        while not os.path.exists(self.stderr_file_path):
            time.sleep(1)
        output = {}
        with open(self.stderr_file_path, 'r', encoding='utf-8') as stderr_file:
            output['stderr'] = stderr_file.read()
        os.remove(self.stderr_file_path)
        output['stdout'] = self._get_stdout()
        return output

    def _handle_magic(self, code: str) -> Dict:
        output = {}
        command_parts = shlex.split(code)
        if not code:
            output['stderr'] = f'No magic specified'
            return output
        if code.lower() == 'reset':
            self.do_shutdown(True)
            return output
        if code.lower() == 'paste':
            return self._handle_paste()
        if command_parts[0] == 'file':
            if len(command_parts) != 2:
                output['stderr'] = 'Usage: %file <file_path>'
            else:
                return self._handle_file_execute(command_parts[1])

        output['stderr'] = f'Invalid magic "{code}"'
        return output

    def _handle_file_execute(self, file_path: str) -> Dict:
        output = {}
        try:
            with open(file_path, 'r') as code_file:
                output.update(self._handle_vbscript_command(code_file.read()))
        except (FileNotFoundError, PermissionError, OSError) as exception:
            output['stderr'] = (''.join(traceback.format_exception(None, exception, None)))
        return output

    def _handle_paste(self) -> Dict:
        output = {}
        win32clipboard.OpenClipboard()
        try:
            clipboard_code = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
            text = clipboard_code
            if not clipboard_code.endswith('\n'):
                text += '\n'
            text += "## -- End pasted text --\n"
            if clipboard_code.lower().strip() != '%paste':
                output.update(self._handle_code(clipboard_code))
            else:
                output['stderr'] = 'Ignored.'
            output['stdout'] = text + output.get('stdout', '')
        except win32clipboard.error as exception:
            output['stderr'] = (''.join(traceback.format_exception(None, exception, None)))
        finally:
            try:
                win32clipboard.CloseClipboard()
            except win32clipboard.error:
                pass
        return output

    def _handle_code(self, code: str) -> Dict:
        output = {}
        if code.strip().lower() in ['exit', 'exit()', 'quit', 'quit()']:
            self._terminate_app()
        elif code.strip().lower() in ['cls', 'clear']:
            os.system('cls')
        elif code.strip().startswith('!'):
            output = self._handle_command_line_code(code.strip()[1:])
        elif code.strip().startswith('%'):
            output = self._handle_magic(code.strip()[1:])
        else:
            output = self._handle_vbscript_command(code)
        return output

    # pylint: disable=too-many-arguments
    def do_execute(self, code, silent, store_history=True, user_expressions=None, allow_stdin=False):

        self.history_manager.append(self.execution_count, code)
        output = self._handle_code(code)

        if not silent:
            if output.get('stdout', list()):
                self.send_response(self.iopub_socket, 'stream', {'name': 'stdout', 'text': output['stdout']})
            if output.get('stderr', list()):
                self.send_response(self.iopub_socket, 'stream',
                                   {'name': 'stderr',
                                    'text': termcolor.colored(output['stderr'], color='red')})
            if output.get('data', list()):
                out_prompt = termcolor.colored(f'Out[{self.execution_count}]:', color='red')
                self.send_response(self.iopub_socket, 'display_data',
                                   {'name': 'stdout',
                                    'data': {'text/plain': f'{out_prompt} {output["data"]}'}})

        return {'status': 'ok', 'execution_count': self.execution_count, 'payload': [], 'user_expressions': {}}

    # pylint: enable=too-many-arguments

    def _is_interpreter_running(self) -> bool:
        return not self.cscript.poll()

    def _shutdown_cleanup(self):
        self.history_manager.disconnect()
        self._send_command('WScript.Quit')
        time.sleep(2)
        if self._is_interpreter_running():
            self.cscript.kill()

    def do_shutdown(self, restart):
        self._shutdown_cleanup()
        if restart:
            self.execution_count = 0
            self.run()
        return {'restart': restart}

    def do_apply(self, content, bufs, msg_id, reply_metadata):
        """DEPRECATED"""
        raise NotImplementedError

    def do_clear(self):
        """DEPRECATED since 4.0.3"""
        raise NotImplementedError

    def do_is_complete(self, code):
        completed = {'status': 'complete'}
        lines = code.splitlines()
        if code.strip().endswith('_'):
            return {'status': 'incomplete'}
        if not any(list(map(str.strip, lines[-1:]))):
            return completed
        for patterns in self.completion_regexes.values():
            if not self._statement_completed(code, **patterns):
                return {'status': 'incomplete', 'indent': self.incomplete_indent}
        return completed

    @staticmethod
    def _statement_completed(code: str, start_pattern: str, end_pattern: str) -> bool:
        flags = re.IGNORECASE
        code_lines = [line.strip() for line in code.splitlines()]
        found_start = any([re.search(start_pattern, line, flags) for line in code_lines])
        found_end = any([re.search(end_pattern, line, flags) for line in code_lines])
        # checks if both None - both start and end not found, or if both found
        return found_start == found_end

    # pylint: disable=too-many-arguments
    def do_history(self, hist_access_type, output, raw, session=None,
                   start=None, stop=None, n=None, pattern=None, unique=False):
        if hist_access_type != "tail" or not n or output:
            return {'history': []}
        result = self.history_manager.tail(n)
        return {'history': result}

    # pylint: enable=too-many-arguments

    def do_complete(self, code, cursor_pos):
        return {'matches': [],
                'cursor_end': cursor_pos,
                'cursor_start': cursor_pos,
                'metadata': {},
                'status': 'ok'}

    def do_inspect(self, code, cursor_pos, detail_level=0):
        return {'status': 'ok', 'data': {}, 'metadata': {}, 'found': False}

    def _terminate_app(self):
        self.cscript.terminate()
        cur_process = psutil.Process()
        parent_process = cur_process.parent()
        parent_process.terminate()
        cur_process.terminate()
