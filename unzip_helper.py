# UNZIP SCRIPT

# Script to UNZIP all Files in a Root Directory (and all of it's children)
# Prints the names of unzipped files.

# Spencer Penn
# 7/19/2017
# Referenced stackoverflow: https://stackoverflow.com/questions/28339000/unzip-zip-files-in-folders-and-subfolders-with-python



import zipfile,fnmatch,os


rootPath = u"//teslamotors.com/US/Finance/Cost IQ/Quotes/M3" # CHOOSE ROOT FOLDER HERE
pattern = '*.zip'
for root, dirs, files in os.walk(rootPath):
    for filename in fnmatch.filter(files, pattern):
        print(os.path.join(root, filename))
        if os.path.join(root, filename).rsplit('\\', 1)[1][0] == '~':
            print('temporary file; not a real file')
            print(os.path.join(root, filename).rsplit('\\', 1)[1][0])
            continue
        elif os.path.join(root, filename).rsplit('\\', 1)[1][0] == '._':
            print('temporary file; not a real file')
            print(os.path.join(root, filename).rsplit('\\', 1)[1][0])
            continue
        else:
            print(os.path.join(root, os.path.splitext(filename)[0]))
            zipfile.ZipFile(os.path.join(root, filename)).extractall(os.path.join(root, os.path.splitext(filename)[0]))
