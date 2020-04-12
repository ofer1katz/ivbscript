from ipykernel.kernelapp import IPKernelApp

from .kernel import VBScriptKernel

IPKernelApp.launch_instance(kernel_class=VBScriptKernel)
