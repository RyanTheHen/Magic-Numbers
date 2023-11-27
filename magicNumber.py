from tkinter import filedialog as di
import pandas as pd
import os

#Define variables to be used later
files = []

#Get current working directory to iterate over
cwd = os.getcwd()

#Defines Magic Numbers by file type for comparison
magic_numbers = {'png': bytes([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]),
                 'jpg': bytes([0xFF, 0xD8, 0xFF, 0xE0]),
                 #*********************#
                 'doc': bytes([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]),
                 'xls': bytes([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]),
                 'ppt': bytes([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]),
                 #*********************#
                 'docx': bytes([0x50, 0x4B, 0x03, 0x04, 0x14, 0x00, 0x06, 0x00]),
                 'xlsx': bytes([0x50, 0x4B, 0x03, 0x04, 0x14, 0x00, 0x06, 0x00]),
                 'pptx': bytes([0x50, 0x4B, 0x03, 0x04, 0x14, 0x00, 0x06, 0x00]),
                 #*********************#
                 'pdf': bytes([0x25, 0x50, 0x44, 0x46]),
                 #*********************#
                 'dll': bytes([0x4D, 0x5A, 0x90, 0x00]),
                 'exe': bytes([0x4D, 0x5A]),

                 }


#Define len for checking header
max_read_size = max(len(m) for m in magic_numbers.values())

 #Compares file to magic numbers to determine file type
def getFileType(file, file_head):
    for ext in magic_numbers:
        if file_head.startswith(magic_numbers[ext]) and file.rsplit('.', 1)[-1] == ext:
            print(f"It's a {ext} File")
            info = [file.rsplit('/', 1)[-1], ext]
            return info
        
        elif file_head.startswith(magic_numbers[ext]):
            if file.rsplit('.', 1)[-1] != ext:
                info = [file.rsplit('/', 1)[-1], f'Found {ext}']
                print(f'Found {ext} instead')
                return info

        elif file.rsplit('.')[-1] not in magic_numbers:
                info = [file.rsplit('/', 1)[-1], 'Not Found']
                return info

#File Scan, opens file in binary and grabs relevant code
def fileScan(files):
    for file in os.listdir(cwd):
        if os.path.isdir(file):
            print('Found Directory')
            files += [[file, "Directory"]]
        else:
            with open(file, 'rb') as fd:
                file_head = fd.read(max_read_size)
            files += [getFileType(fd.name, file_head)]
    files = fixNone(files)
    return files


def getIndex(files):
    nums = []
    x = 0
    for i in files:
        x += 1
        
        nums += [f'File {x}']
    return nums

#Fixing bug where getFileType sometimes returns None
def fixNone(files):
    for item in files:
        if item == None:
            files.remove(None)
            return files
        else:
            for i in item:
                    if i==None:
                        files.remove(None)
                        return files

#Write file information to Excel sheet
def writeData(files, nums):
    with pd.ExcelWriter('output.xlsx', if_sheet_exists='new', mode='a') as writer:
        
        array = pd.DataFrame(files,
            columns=['File Name', 'File Type'],
            index=[nums])
        array.to_excel(writer, sheet_name='Data')

files += fileScan(files)
files = fixNone(files)
nums = getIndex(files)
files = fixNone(files)

#Fixing where nums ends uo with an extra instance and prevents writing to sheet
nums.pop()

#Output to console for error checking
print(files)
print(nums)

#Call Function to Write to Excel
writeData(files, nums)