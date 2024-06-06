import os
import shutil

def createDir(path):
    try:
        os.mkdir(path)
    except FileExistsError:
        pass

def createFolders():
    try:
        os.chdir("C:\\Users\\husaa\\Desktop")
    except FileNotFoundError as e:
        print(e)

    createDir("Programs")

    try:
        os.chdir("C:\\Users\\husaa\\Documents")
    except FileNotFoundError as e:
        print(e)

    createDir("Text Files")
    createDir("PDF Documents")
    createDir("Spreadsheets")
    createDir("Word Documents")

def checkDesktop():
    os.chdir("C:\\Users\\husaa\\Desktop")
    files = os.listdir()

    for filename in files:
        filename = filename.split('.')
        if ( len(filename) == 1 ): # This means its a folder, not a files
            continue

        else:
            if ( filename[1] == "txt" ):
                filename = '.'.join(filename)
                shutil.move(f"C:\\Users\\husaa\\Desktop\\{filename}",
                            f"C:\\Users\\husaa\\Documents\\Text Files\\{filename}")

def main():
    createFolders()
    checkDesktop()

main()
