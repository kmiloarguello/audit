from cx_Freeze import setup,Executable

includes = []
excludes = []
packages = []
filename = "App.py"

setup(
    name = 'Auditorias',
    version = '0.1',
    description = 'Auditorias',
    author = 'KAPTA',
    author_email = 'camilo.arguello@kapta.biz',
    options = {'build_exe': {'excludes':excludes,'packages':packages,'includes':includes}},
    executables = [Executable(filename, base = "Win32GUI", icon = None)])
