# iVBScript

Interactive VBScript

Jupyter kernel implementation for VBScript

# Installation*
```shell script
git clone https://github.com/ofer1katz/ivbscript.git 
cd ivbscript
git lfs pull 
python setup.py develop
```
### **notice:*
*during installation registry will be edited to allow ANSI console color*

```shell script
[HKEY_CURRENT_USER\Console]
"VirtualTerminalLevel"=dword:00000001
```

*during installation com dll "tlbinf32.dll" (TLI.TLIApplication) will be registered if not registered already.*

*%systemroot%\SysWOW64\regsvr32.exe can be used to manually register/unregister*

---

*kernel specs will be copied to `os.path.join(os.path.abspath(sys.prefix), 'share', 'jupyter', 'kernels', 'vbscript')`*

# Usage
```shell script
ivbscript
```
or
```shell script
jupyter console --kernel vbscript
```

#### Special Commands
- `cls/clear` - clear console
- `exit/exit()/quit/quit()` - exit iVBScript
- `!<command>` - execute a child program in a new process
- `<variable>?` - inspect `<variable>`
- `%reset` - reset console
- `%file <file_path>` - read `<file_path>` and run the content as VBScript code
- `%paste` - paste and execute

####If you are having this error:
```python
Unhandled exception in event loop:
  File "c:\users\USER\appdata\local\programs\python\python38-32\lib\asyncio\proactor_events.py", line 768, in _loop_self_reading
    f.result()  # may raise
  File "c:\users\USER\appdata\local\programs\python\python38-32\lib\asyncio\windows_events.py", line 808, in _poll
    value = callback(transferred, key, ov)
  File "c:\users\USER\appdata\local\programs\python\python38-32\lib\asyncio\windows_events.py", line 457, in finish_recv
    raise ConnectionResetError(*exc.args)

Exception [WinError 995] The I/O operation has been aborted because of either a thread exit or an application request
Press ENTER to continue...
```
Take a look at the following workaround:
https://github.com/ipython/ipython/issues/12049#issuecomment-586544339

# Development
### Install development requirements
```shell script
pip install -r requirements_dev.txt -U --upgrade-strategy eager
```

#### Tests
```shell script
coverage erase
coverage run --source=. --omit="*\tests\*" -m pytest -v -s
coverage report -m
```

#### Code Analytics
```shell script
prospector --strictness veryhigh
# Analyze the given Python modules and compute Cyclomatic Complexity (CC).
radon cc . --min B
# Analyze the given Python modules and compute the Maintainability Index.
radon mi . --min B
# Analyze the given Python modules and compute raw metrics.
radon raw .
# Analyze the given Python modules and compute their Halstead metrics.
radon hal .
```

### TODO:
- [ ] test coverage
- [ ] using pipes instead of files for communication with vbscript
- [ ] better implementation of exit/quit (via jupyter)
- [ ] evaluate expressions
- [ ] completion using - `Tab` (`do_complete`)
- [ ] use `do_inspect()` for inspect
- [ ] inspect levels support - `?`/`??`
- [ ] paste into terminal - `Ctrl + v`
