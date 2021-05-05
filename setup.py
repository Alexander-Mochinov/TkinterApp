from cx_Freeze import setup, Executable

base = None    

executables = [Executable("interface.py", base=base)]

packages = ["tkinter", "idna", "openpyxl", "os", "shutil", "re", "sys", "smtplib"]
options = {
    'build_exe': {    
        'packages':packages,
    },    
}

setup(
    name = "send mail",
    options = options,
    version = "1.0",
    description = 'Sending mail',
    executables = executables
)