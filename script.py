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
import logging

def setup_logging:
    pass

def getLoggedOnUser():
    try:
        c = wmi.WMI()
        for session in c.Win32_LogonSession():
            if session.LogonType == 2:  # 2 is Interactive Logon
                users = session.references("Win32_LoggedOnUser")
                if users:
                    return users[0].Antecedent.Name
    except Exception as e:
        #        logging.info(traceback.format_exc())
#        logging.error("Error in get logged on user")
        pass

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
        new_dir = ""
        match ext:
            case "jpeg" | "jpg" | "gif" | "png" | "tiff":
                shutil.move(f"{dir}{filename}", f"{images_dir}{filename}")
                new_dir = images_dir

            case "txt":
                shutil.move(f"{dir}{filename}", f"{txt_dir}{filename}")
                new_dir = txt_dir

            case "pdf" | "PDF":
                shutil.move(f"{dir}{filename}", f"{pdf_dir}{filename}")
                new_dir = pdf_dir

            case "xls" | "xlsx" | "xlsm":
                shutil.move(f"{dir}{filename}", f"{spreadsheet_dir}{filename}")
                new_dir = spreadsheet_dir

            case "doc" | "docx":
                shutil.move(f"{dir}{filename}", f"{word_dir}{filename}")
                new_dir = word_dir

            case "xml":
                shutil.move(f"{dir}{filename}", f"{xml_dir}{filename}")
                new_dir = xml_dir

            case "exe" | "lnk":
                if ( dir != desktop_dir ):
                    if ( "INSTALL" not in filename.upper() and "SETUP" not in filename.upper() ):
                        shutil.move(f"{dir}{filename}", f"{desktop_dir}{filename}")
                        new_dir = desktop_dir
        if new_dir:
            #            logging.info(f"Moved file {filename}.{ext} from {dir} to {new_dir}")
            pass

    except Exception as e:
        #        logging.error(str(e))
        pass

def createDir(path):
    try:
        os.makedirs(path, exist_ok=True)
    except Exception as e:
        #        logging.error(str(e))
        pass

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
        #        logging.error(str(e))
        pass

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
        #        logging.error(str(e))
        pass

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
                    #                    logging.info("Stop request recieved")
                    break
        except Exception as e:
            #            logging.info(traceback.format_exc())
            pass

    def SvcStop(self):
        self.ReportServiceStatus(win_svc.SERVICE_STOP_PENDING)
#        logging.info("Service stopping")
        win_event.SetEvent(self.event)
        self.obs1.stop()
        self.obs2.stop()
        self.obs1.join()
        self.obs2.join()
        self.ReportServiceStatus(win_svc.SERVICE_STOPPED)

if __name__ == "__main__":
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(FileOrganizerService)
        servicemanager.StartServiceCtrlDispatcher()

    else:
        win_svc_util.HandleCommandLine(FileOrganizerService)
