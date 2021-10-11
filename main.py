import os
import zipfile
import py7zr
from pandas import read_excel


def unzip_test():


    # dictionary zID-Group
    file_name = 'Class-List.xlsx'
    my_sheet = 'Sheet1'  # change it to your sheet name, you can find your sheet name at the bottom left of your excel file
    df = read_excel(file_name, sheet_name=my_sheet, header=None)
    print(df.shape)

    zid2group = dict()
    for i in range(85):
        zid2group[str(df.iloc[i, 0])] = df.iloc[i, 5]

    # print(zid2group)



    path = os.getcwd()
    path1 = os.path.join(path, 'ZEIT1501-5217_00166_CATIA Test Submission Box_submission')

    all_folders = os.listdir(path1)
    a = 1
    for folder in all_folders:
        print(folder)
        items = folder.split('_')
        zid = items[0]
        print(zid)
        a = a+1
        outfolder = os.path.join(os.getcwd(), 'submissions')
        outfolder2 = outfolder





        path2 = os.path.join(path1, folder)

        filename = os.listdir(path2)
        print(filename)
        if len(filename) > 1:
            print(path2 + 'has more files')


        for filenamei  in filename:
            checkzip = filenamei.split('.')
            if checkzip[-1] not in ['zip', '7z', 'rar']:
                continue
            zipname  = os.path.join(path2, filenamei)
            # handle = zipfile.ZipFile(zipname)
            # handle.extractall(path)
            path3 = os.path.join(outfolder2, zid)
            try:
                os.mkdir(path3)
            except FileExistsError:
                pass
            if checkzip[-1] in ['7z']:
                with py7zr.SevenZipFile(zipname, 'r') as archive:
                    archive.extractall(path=path3)
            if checkzip[-1] in ['zip']:
                handle = zipfile.ZipFile(zipname)
                handle.extractall(path3)
    print(a)

def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    unzip_test()


