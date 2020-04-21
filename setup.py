import ctypes
import os
import struct
import sys
import traceback
import winreg

import termcolor
from jupyter_client.kernelspec import KernelSpecManager, NoSuchKernel
from setuptools import find_packages, setup

from ivbscript import __version__
from ivbscript.kernel import VBScriptKernel


class RegistryWrongValue(Exception):
    pass


def allow_ansi_console_color_if_needed():
    root = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
    key = winreg.OpenKey(root, 'Console')
    value_name = 'VirtualTerminalLevel'
    value_type = winreg.REG_DWORD
    value = 1
    try:
        try:
            reg_value, reg_type = winreg.QueryValueEx(key, value_name)
            if reg_value != value or reg_type != value_type:
                winreg.DeleteValue(key, value_name)
                raise RegistryWrongValue
        except (FileNotFoundError, RegistryWrongValue):
            winreg.SetValueEx(key, value_name, None, value_type, value)

    except PermissionError as exception:
        print(''.join(traceback.format_exception(None, exception, None)))
        print('Allow ansi console color manually')


def is_tlbinf32_registered() -> bool:
    root = winreg.ConnectRegistry(None, winreg.HKEY_CLASSES_ROOT)
    try:
        winreg.OpenKey(root, "TLI.TLIApplication")
        return True
    except FileNotFoundError:
        return False
    finally:
        root.Close()


def register_tlbinf32():
    bits = struct.calcsize('P') * 8
    if bits != 32:
        print(termcolor.colored('Python is 64 bit. Register tlbinf32 manually', color='red'))
    else:
        try:
            dll_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'tlbinf32.dll')
            dll = ctypes.OleDLL(dll_path)
            dll.DllRegisterServer()
        except OSError as exception:
            print(termcolor.colored(''.join(traceback.format_exception(None, exception, None)), color='red'))
            print(termcolor.colored('check git lfs', color='red'))
            raise exception


def register_tlbinf32_if_needed():
    if not is_tlbinf32_registered():
        register_tlbinf32()
        print('tlbinf32 Installed')
    else:
        print('tlbinf32 Already Installed')


def is_kernel_spec_installed() -> bool:
    try:
        KernelSpecManager().get_kernel_spec(VBScriptKernel.language)
        return True
    except NoSuchKernel:
        return False


def install_kernel_spec():
    default_spec_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'defaultspec')
    KernelSpecManager().install_kernel_spec(default_spec_dir, VBScriptKernel.language, prefix=sys.prefix)


def install_kernel_spec_if_needed():
    if not is_kernel_spec_installed():
        install_kernel_spec()
        print('Kernel Installed')
    else:
        print('Kernel Already Installed')


allow_ansi_console_color_if_needed()
register_tlbinf32_if_needed()
install_kernel_spec_if_needed()

with open('requirements.txt', 'r') as f:
    REQUIREMENTS = f.readlines()

setup(
    name="ivbscript",
    version=__version__,
    packages=find_packages(),
    description="VBScript Jupyter Kernel",
    classifiers=[
        'Programming Language :: VBScript',
        "Programming Language :: Python :: 3"
        "Operating System :: Windows"
    ],
    entry_points={
        'console_scripts': ['ivbscript=ivbscript.app:main']
    },
    install_requires=REQUIREMENTS,
    python_requires='>=3.8',
)
