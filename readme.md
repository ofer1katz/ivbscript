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

*during installation com dll "tlbinf32.dll" (TLI.TLIApplication) will be registered if not registered already.*

*regsvr32.exe can be used to manually register/unregister*

---

*kernel specs will be copied to `os.path.join(os.path.abspath(sys.prefix), 'share', 'jupyter', 'kernels', 'vbscript')`*

# Usage
```shell script
jupyter console --kernel vbscript
```

# Development
### Install development requirements
```shell script
pip install -r requirements_dev.txt -U --upgrade-strategy eager
```

#### Tests
```shell script
coverage erase
coverage run --source=. --omit="*\tests\*" -m pytest
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
- [ ] create launch shortcut instead of `jupyter console --kernel vbscript`
- [ ] completion using - `Tab` (`do_complete`)
- [ ] use `do_inspect()` for inspect
- [ ] inspect levels support - `?`\\`??` 
- [ ] paste into terminal - `%paste`
- [ ] clear terminal - `cls`\\`clear`
