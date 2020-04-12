import ctypes
import os
import sys
import traceback
import winreg

import termcolor
from jupyter_client.kernelspec import KernelSpecManager, NoSuchKernel
from setuptools import find_packages, setup

from ivbscript import __version__
from ivbscript.kernel import VBScriptKernel


def is_tlbinf32_installed():
    root = winreg.ConnectRegistry(None, winreg.HKEY_CLASSES_ROOT)
    try:
        winreg.OpenKey(root, "TLI.TLIApplication")
        return True
    except FileNotFoundError:
        return False
    finally:
        root.Close()


def install_tlbinf32():
    try:
        dll = ctypes.OleDLL('tlbinf32.dll')
        dll.DllRegisterServer()
    except OSError as exception:
        print(termcolor.colored(''.join(traceback.format_exception(None, exception, None)), color='red'))
        print(termcolor.colored('check git lfs', color='red'))
        raise exception


def install_tlbinf32_if_needed():
    if not is_tlbinf32_installed():
        install_tlbinf32()
        print('tlbinf32 Installed')
    else:
        print('tlbinf32 Already Installed')


def is_kernel_spec_installed():
    try:
        KernelSpecManager().get_kernel_spec(VBScriptKernel.implementation)
        return True
    except NoSuchKernel:
        return False


def install_kernel_spec():
    default_spec_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'defaultspec')
    KernelSpecManager().install_kernel_spec(default_spec_dir, VBScriptKernel.implementation, prefix=sys.prefix)


def install_kernel_spec_if_needed():
    if not is_kernel_spec_installed():
        install_kernel_spec()
        print('Kernel Installed')
    else:
        print('Kernel Already Installed')


install_tlbinf32_if_needed()
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
        'console_scripts': []
    },
    install_requires=REQUIREMENTS,
    python_requires='>=3.8',
)
