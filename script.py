import os
import shutil

desktop_dir = "C:\\Users\\husaa\\Desktop\\"
downloads_dir = "C:\\Users\\husaa\\Downloads\\"
documents_dir = "C:\\Users\\husaa\\Documents\\"
images_dir = "C:\\Users\\husaa\\Gallery\\"
txt_dir = "C:\\Users\\husaa\\Documents\\Text Files\\"
pdf_dir = "C:\\Users\\husaa\\Documents\\PDF Documents\\"
spreadsheet_dir = "C:\\Users\\husaa\\Documents\\Spreadsheets\\"
word_dir = "C:\\Users\\husaa\\Documents\\Word Documents\\"
xml_dir = "C:\\Users\\husaa\\Documents\\XML Documents\\"


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
        filename = filename.split('.')
        if ( len(filename) == 1 ): # This means its a folder, not a file
            continue

        else:
            ext = filename[1]
            filename = '.'.join(filename)

            match ext:
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
                        shutil.move(f"{dir}{filename}", f"{desktop_dir}{filename}")

def main():
    createFolders()
    while(True):
        checkDir(desktop_dir)
        checkDir(downloads_dir)

main()
