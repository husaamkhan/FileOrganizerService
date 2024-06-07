import os
import shutil
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

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

class EventHandler(FileSystemEventHandler):
    def on_created(self, event):
        pass
#        if not os.path.isdir(event.src_path):
#            dir, filename = os.path.split(event.src_path)
#            dir = f"{dir}\\"
#            ext = filename.split('.')[1]
#            print(event.src_path)

#            if ( ext not in ["tmp", "crdownload", "part"] ): # That means a file was downloaded and the temp file was detected
#                print(filename)
#                moveFile(dir, filename, ext)

    def on_modified(self, event):
        if not os.path.isdir(event.src_path):
            print(f"Modified: {event.src_path}")
            dir, filename = os.path.split(event.src_path)
            dir = f"{dir}\\"
            ext = filename.split('.')[1]
            print(ext)

            if ( ext != "tmp" and ".crdownload" not in filename ):
                print(f"Not temp: {filename}")
                if ( os.path.exists(event.src_path) ):
                    print("file still exists, moving file")
                    moveFile(dir, filename, ext)

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

def main():
    print("Husaam's file organizer is running!")
    createFolders()
#    while(True):
#        checkDir(desktop_dir)
#        checkDir(downloads_dir)

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

main()
