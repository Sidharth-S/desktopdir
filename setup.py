from cx_Freeze import setup, Executable
import os
import shutil
from pathlib import Path
import time
import winreg as reg

# add func to kill desktopdir.exe if aldready running 

build_exe_options = {'packages':['os', 'winshell', 'win32com','win32api','time'], #pacakges used
                     'excludes':['tkinter','numpy','matplotlib','scipy'],
                     "include_msvcr": True,
                     'include_files':['startddir.vbs'],
                     "build_exe" : "C:\\Program Files\\Prosid\\desktopdir"}


path = "C:\\ProgramData\\Prosid\\desktopdir"
if os.path.isdir("C:\\Program Files\\Prosid\\desktopdir"):shutil.rmtree("C:\\Program Files\\Prosid\\desktopdir")


executables = [Executable(
            "desktopdir.py",
            target_name = "desktopdir",
            copyright="Copyright (C) 2022 Prosid",
            base=None,)]

setup(
    name="desktopdir",
    version="1.0",
    author = "Sidharth S",
    description="Drive shortcut creator in desktop for Windows",
    executables=executables,
    options={
        "build_exe": build_exe_options,
    },
)

time.sleep(3)

startup_folder = os.path.join(Path.home(),r"AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Startup\\" )

address = "C:\\Program Files\\Prosid\\desktopdir\\startddir.vbs"

key = reg.HKEY_CURRENT_USER
key_value = "Software\Microsoft\Windows\CurrentVersion\Run"
    
open = reg.OpenKey(key,key_value,0,reg.KEY_ALL_ACCESS)
    
reg.SetValueEx(open,"desktopdir",0,reg.REG_SZ,address)
    
reg.CloseKey(open)

print("Installation complete !")