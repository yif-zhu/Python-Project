import os
from cx_Freeze import setup, Executable

os.environ['TCL_LIBRARY'] = 'C:\\ProgramData\\Anaconda3\\tcl\\tcl8'
os.environ['TK_LIBRARY'] = 'C:\\ProgramData\\Anaconda3\\tcl\\tk8.6'

# Dependencies are automatically detected, but it might need
# fine tuning.
buildOptions = dict(
    packages = [],
    excludes = [],
    include_files=['C:\\ProgramData\\Anaconda3\\DLLs\\tcl86t.dll', 'C:\\ProgramData\\Anaconda3\\DLLs\\tk86t.dll']
)

import sys
base = 'Win32GUI' if sys.platform=='win32' else None

executables = [
    Executable('C:\\PyCharm\\untitled1\\QuickFillExcel.py', base=base)
]

setup(name='editor',
      version = '1.0',
      description = '',
      options = dict(build_exe = buildOptions),
      executables = executables)