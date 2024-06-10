import os
import shutil
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import win32serviceutil as win_svc_util
import win32service as win_svc
import win32event as win_event
import servicemanager
import sys
import wmi
import ctypes

def showErrorMessage(message):
    WMI_OK = 0x00000000
    WMI_ICON = 0x00000010
    ctypes.windll.user32.MessageBoxW(0, message, "FileOrganizerService Error", WMI_OK| WMI_ICON)

def getLoggedOnUser():
    try:
        c = wmi.WMI()
        for session in c.Win32_LogonSession():
            if session.LogonType == 2:  # 2 is Interactive Logon
                users = session.references("Win32_LoggedOnUser")
                if users:
                    return users[0].Antecedent.Name
    except Exception as e:
        showErrorMessage(f"Error identifying logged in user: {e}")

    return None

def setDirs(user):
    global home_dir, desktop_dir, downloads_dir, documents_dir, images_dir, txt_dir, pdf_dir
    global spreadsheet_dir, word_dir, xml_dir

    home_dir = f"C:\\Users\\{user}"
    desktop_dir = f"{home_dir}\\Desktop\\"
    downloads_dir = f"{home_dir}\\Downloads\\"
    documents_dir = f"{home_dir}\\Documents\\"
    images_dir = f"{home_dir}\\Pictures\\"
    txt_dir = f"{documents_dir}Text Files\\"
    pdf_dir = f"{documents_dir}PDF Documents\\"
    spreadsheet_dir = f"{documents_dir}Spreadsheets\\"
    word_dir = f"{documents_dir}Word Documents\\"
    xml_dir = f"{documents_dir}XML Documents\\"

def moveFile(dir, filename, ext):
    try:
        match ext:
            case "jpeg" | "jpg" | "gif" | "png" | "tiff":
                shutil.move(f"{dir}{filename}", f"{images_dir}{filename}")

            case "txt":
                shutil.move(f"{dir}{filename}", f"{txt_dir}{filename}")

            case "pdf" | "PDF":
                shutil.move(f"{dir}{filename}", f"{pdf_dir}{filename}")

            case "xls" | "xlsx" | "xlsm":
                shutil.move(f"{dir}{filename}", f"{spreadsheet_dir}{filename}")

            case "doc" | "docx":
                shutil.move(f"{dir}{filename}", f"{word_dir}{filename}")

            case "xml":
                shutil.move(f"{dir}{filename}", f"{xml_dir}{filename}")

            case "exe" | "lnk":
                if ( dir != desktop_dir ):
                    if ( "INSTALL" not in filename.upper() and "SETUP" not in filename.upper() ):
                        shutil.move(f"{dir}{filename}", f"{desktop_dir}{filename}")
    except Exception as e:
        showErrorMessage(f"Error moving file {filename}.{ext}: {e}")

def createDir(path):
    try:
        os.makedirs(path, exist_ok=True)
    except Exception as e:
        showErrorMessage(f"Error creating directory {path}: {e}")

def createFolders():
    try:
        os.chdir(desktop_dir)

        createDir("Programs")

        os.chdir(documents_dir)

        createDir("Text Files")
        createDir("PDF Documents")
        createDir("Spreadsheets")
        createDir("Word Documents")
        createDir("XML Documents")
    except Exception as e:
        showErrorMessage(f"Error creating folders: {e}")

def checkDir(dir):
    try:
        os.chdir(dir)
        files = os.listdir()

        for filename in files:
            if os.path.isdir(filename):
                continue

            ext = filename.split('.')[1]
            moveFile(dir, filename, ext)
    except Exception as e:
        showErrorMessage(f"Error checking directories: {e}")

class EventHandler(FileSystemEventHandler):
    def on_modified(self, event):
        if not os.path.isdir(event.src_path):
            dir, filename = os.path.split(event.src_path)
            dir = f"{dir}\\"
            ext = filename.split('.')[1]
            
            if ( ext != "tmp" and ".crdownload" not in filename ):
                if ( os.path.exists(event.src_path) ):
                    moveFile(dir, filename, ext)

class FileOrganizerService(win_svc_util.ServiceFramework):
    _svc_name_ = "FileOrganizerService"
    _svc_display_name_ = "FileOrganizerService"
    _svc_description_ = "Automatically organizes Desktop and Downloads files by filetype."

    event_handler = EventHandler() 
    obs1 = Observer()

    obs2 = Observer()

    def __init__(self, args):
        win_svc_util.ServiceFramework.__init__(self, args)
        self.event = win_event.CreateEvent(None, 0, 0, None)

    def GetAcceptedControls(self):
        result = win_svc_util.ServiceFramework.GetAcceptedControls(self)
        result |= win_svc.SERVICE_ACCEPT_PRESHUTDOWN
        return result

    def SvcDoRun(self):
        try:
            self.ReportServiceStatus(win_svc.SERVICE_RUNNING)

            user = getLoggedOnUser()
            setDirs(user)

            createFolders()
            checkDir(desktop_dir)
            checkDir(downloads_dir)

            self.obs1.schedule(self.event_handler, path=desktop_dir, recursive=False)
            self.obs2.schedule(self.event_handler, path=downloads_dir, recursive=False)

            self.obs1.start()
            self.obs2.start()

            while True:
                code = win_event.WaitForSingleObject(self.event, win_event.INFINITE)
                if ( code == win_event.WAIT_OBJECT_0 ):
                    break
        except Exception as e:
            showErrorMessage(f"Error running service: {e}")

    def SvcStop(self):
        self.ReportServiceStatus(win_svc.SERVICE_STOP_PENDING)
        win_event.SetEvent(self.event)
        self.obs1.stop()
        self.obs2.stop()
        self.obs1.join()
        self.obs2.join()
        self.ReportServiceStatus(win_svc.SERVICE_STOPPED)

if __name__ == "__main__":
    try:
        if len(sys.argv) == 1:
            servicemanager.Initialize()
            servicemanager.PrepareToHostSingle(FileOrganizerService)
            servicemanager.StartServiceCtrlDispatcher()

        else:
            win_svc_util.HandleCommandLine(FileOrganizerService)
    except Exception as e:
        showErrorMessage(f"Error starting service: {e}")
