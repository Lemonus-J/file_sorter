import os
import win32com.client
import traceback

def refresh_windows_explorer():
    shell = win32com.client.Dispatch("Shell.Application")
    for window in shell.Windows():
        if window.Name == 'File Explorer':
            window.Refresh()

def sort_directory():
    filesInDir = os.listdir()
    currScript = os.path.basename(__file__)
    currWorkingDir = os.getcwd()

    for file in filesInDir: 

        if (os.path.isdir(file)):
            continue

        try:
            splitFileName = file.split('.', -1)

            fileName = splitFileName[0]
            fileType = splitFileName[-1]

            # skip if file is this script file
            if (file == currScript or file == f'{fileName}.exe'):
                continue
        except:
            traceback.print_exc()
            continue

        # create dir of fileType if it does not exist
        if (filesInDir.__contains__(fileType) == False):
            os.mkdir(fileType)
            filesInDir = os.listdir()

        os.replace(f'{currWorkingDir}/{file}', f'{currWorkingDir}/{fileType}/{file}')

    refresh_windows_explorer()

sort_directory()

