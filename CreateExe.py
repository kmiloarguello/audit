from distutils.core import setup
import py2exe, sys, os

sys.argv.append('py2exe')

setup(
    options = {'py2exe': {'bundle_files': 3, 'compressed': True, 'dll_excludes': ['tcl85.dll', 'tk85.dll']}},
    console = [{'script': 'App.py'}],
    zipfile = None,
    )