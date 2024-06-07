import os
import shutil
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import win32serviceutil as win_service
import win32event as win_event

desktop_dir = "C:\\Users\\husaa\\Desktop\\"
downloads_dir = "C:\\Users\\husaa\\Downloads\\"
documents_dir = "C:\\Users\\husaa\\Documents\\"
images_dir = "C:\\Users\\husaa\\Pictures\\"
txt_dir = "C:\\Users\\husaa\\Documents\\Text Files\\"
pdf_dir = "C:\\Users\\husaa\\Documents\\PDF Documents\\"
spreadsheet_dir = "C:\\Users\\husaa\\Documents\\Spreadsheets\\"
word_dir = "C:\\Users\\husaa\\Documents\\Word Documents\\"
xml_dir = "C:\\Users\\husaa\\Documents\\XML Documents\\"

def moveFile(dir, filename, ext):
    match ext:
        case "jpeg" | "jpg" | "gif" | "png" | "tiff":
            shutil.move(f"{dir}{filename}", f"{images_dir}{filename}")

        case "txt":
            shutil.move(f"{dir}{filename}", f"{txt_dir}{filename}")

        case "pdf" | "PDF":
            shutil.move(f"{dir}{filename}", f"{pdf_dir}{filename}")

        case "xls" | "xlsx" | "xlsm" | "csv":
            shutil.move(f"{dir}{filename}", f"{spreadsheet_dir}{filename}")

        case "doc" | "docx":
            shutil.move(f"{dir}{filename}", f"{word_dir}{filename}")

        case "xml":
            shutil.move(f"{dir}{filename}", f"{xml_dir}{filename}")

        case "exe" | "lnk":
            if ( dir != desktop_dir ):
                if ( "INSTALL" not in filename.upper() and "SETUP" not in filename.upper() ):
                    shutil.move(f"{dir}{filename}", f"{desktop_dir}{filename}")


def createDir(path):
    try:
        os.mkdir(path)
    except FileExistsError:
        pass

def createFolders():
    try:
        os.chdir(desktop_dir)
    except FileNotFoundError as e:
        print(e)

    createDir("Programs")

    try:
        os.chdir(documents_dir)
    except FileNotFoundError as e:
        print(e)

    createDir("Text Files")
    createDir("PDF Documents")
    createDir("Spreadsheets")
    createDir("Word Documents")
    createDir("XML Documents")

def checkDir(dir):
    os.chdir(dir)
    files = os.listdir()

    for filename in files:
        if os.path.isdir(filename):
            continue

        ext = filename.split('.')[1]
        moveFile(dir, filename, ext)

class EventHandler(FileSystemEventHandler):
    def on_modified(self, event):
        if not os.path.isdir(event.src_path):
            dir, filename = os.path.split(event.src_path)
            dir = f"{dir}\\"
            ext = filename.split('.')[1]

            if ( ext != "tmp" and ".crdownload" not in filename ):
                print(f"Not temp: {filename}")
                if ( os.path.exists(event.src_path) ):
                    print("file still exists, moving file")
                    moveFile(dir, filename, ext)

class FileOrganizerService(win_service.ServiceFramework):
    _scv_name_ = "FileOrganizerService"
    _svc_display_name_ = "FileOrganizerService"
    _svc_description_ = "Automatically organizes Desktop and Downloads files by filetype."

    def __init__(self, args):
        win_service.ServiceFramework.__init__(self, args)
        self.event = win_event.CreateEvent(None, 0, 0, None)

    def GetAcceptedControls(self):
        result = win_service.ServiceFramework.GetAcceptedControls(self)
        result |= win_service.SERVICE_ACCEPT_PRESHUTDOWN
        return result

    def SvcDoRun(self):
        run()

    def SvcStop(self):
        self.ReportServiceStatus(win_service.SERVICE_STOP_PENDING)
        win_event.SetEvent(self.event)

def run():
    print("Husaam's file organizer is running!")
    createFolders()
    checkDir(desktop_dir)
    checkDir(downloads_dir)

    event_handler = EventHandler() 
    obs1 = Observer()
    obs1.schedule(event_handler, path=desktop_dir, recursive=False)

    obs2 = Observer()
    obs2.schedule(event_handler, path=downloads_dir, recursive=False)

    obs1.start()
    obs2.start()
    
    while True:
        try:
            pass
        except KeyboardInterrupt as e:
            print(e)
            obs1.stop()
            obs2.stop()
            print("Terminating...")

if __name__ == "__main__":
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(FileOrganizerService)
        servicemanager.StartServiceCtrlDispatcher()

    else:
        win_service.HandleCommandLine(FileOrganizerService)
