import os, winshell, win32com.client,win32api
import time

def createshortcut(target,location,shortcut_name):

    path = os.path.join(location, f'{shortcut_name}.lnk')

    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortCut(path)
    shortcut.Targetpath = target
    shortcut.IconLocation = target
    
    shortcut.save()
    return

  
def deskdir():
    desktop = winshell.desktop() #obtain user's desktop path
    
    #list containing current created drive shortcuts as full path ["C:\\Users\\Name\\OneDrive\\Desktop\\C.lnk","C:\\Users\\Name\\OneDrive\\Desktop\\E.lnk"]
    created = list() 

    #list containing drives used for previous creation loop ['C:\\','E:\\','G:\\']
    existingdrives = list()

    while True:
        time.sleep(1)

        drives = win32api.GetLogicalDriveStrings()
        drives = drives.split('\000')[:-1]  # ['C:\\','E:\\','G:\\']
        
        # True if any chnage in drive structure
        if existingdrives != drives:
            # Remove all created shortcuts
            for i in created:
                try:os.remove(i)
                except: pass

            #create new shortcuts
            for i in drives: 
                createshortcut(i,desktop,i[0])
        

            dirlist = [os.path.join(desktop,f"{x[0]}.lnk") for x in drives]
            # ["C:\\Users\\Name\\OneDrive\\Desktop\\C.lnk","C:\\Users\\Name\\OneDrive\\Desktop\\E.lnk"]

            created = created + dirlist
            created = list(set(created))
            existingdrives = drives
    



deskdir()
