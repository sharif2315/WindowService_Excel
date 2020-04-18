import os
import pyodbc
import shutil

mypath = r'C:\Users\sahmed243\Documents\WebDevelopment\Python\Scripts\excel_Import\test'
myarchive = r'C:\Users\sahmed243\Documents\WebDevelopment\Python\Scripts\excel_Import\test\archive'
filelist = []

# reading directory
for dirpath, dirnames, filenames in os.walk(mypath):

    for filename in filenames:

        # Reading directory for only xlsx files which are in dir test
        if (filename.lower().endswith('.xlsx')) and (dirpath == mypath):
            filepath = dirpath + '\\' + filename
            filelist.append(filepath)
        else:
            continue

for file in filelist:
    filename = os.path.basename(file)

    shutil.move( file, str(myarchive + '\\' + filename) )
    #print('File: ' + str(file) + ' has been successfully moved to archive folder')
    print('File: ' + str(filename) + ' has been successfully moved to archive folder')
