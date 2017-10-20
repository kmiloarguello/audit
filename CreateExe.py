from cx_Freeze import setup,Executable

includes = []
excludes = []
packages = []
filename = "App.py"

exe = Executable(
   script = "App.py",
   base = 'Win32GUI',
   targetName = "AUDITORIAS.exe",
   icon = 'kapta_mex.ico'
)

setup(
    name = 'Auditorias',
    version = '0.1',
    description = 'Auditorias',
    author = 'KAPTA',
    author_email = 'camilo.arguello@kapta.biz',
    options = {'build_exe': {'excludes':excludes,'packages':packages,'includes':includes}},
    executables = [exe])
