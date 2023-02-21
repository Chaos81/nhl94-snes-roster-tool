from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need
# fine tuning.

import sys, os

packages = []
PYTHON_INSTALL_DIR = os.path.dirname(os.path.dirname(os.__file__))
buildOptions = dict(packages = [], excludes = [], include_files = ['nhl94.gif', 'icon.ico',
            os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tk86t.dll'),
            os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tcl86t.dll')])
base = 'Win32GUI'

executables = [
    Executable('SNES Roster Tool.py', base=base, icon="icon.ico")
]

os.environ['TCL_LIBRARY'] = r'C:\Users\John\AppData\Local\Programs\Python\Python36-32\tcl\tcl8.6'
os.environ['TK_LIBRARY'] = r'C:\Users\John\AppData\Local\Programs\Python\Python36-32\tcl\tk8.6'

setup(name='SNES NHL 94 Roster Tool',
      version = '0.6',
      description = 'Extract/Import rosters and player attributes from/to SNES NHL 94 ROMs',
      options = dict(build_exe = buildOptions),
      executables = executables)
