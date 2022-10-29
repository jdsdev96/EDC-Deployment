from codecs import raw_unicode_escape_decode
import sys
from subprocess import check_call, check_output
import os
import csv
import shutil
print("\u001b[4m\u001b[35mLayout Import Script\u001b[0m")
print("\u001b[37m\u001b[0mPython Version " + sys.version)
if sys.version[:7] != "3.10.8 ":
    print("\u001b[33m***The version of Python is different from what this script was written on. Errors may occur.***")


def inst_openpyxl():
    print("\u001b[33m\n Installing openpyxl...")
    # implement pip as a subprocess:
    check_call([sys.executable, '-m', 'pip', 'install', 'openpyxl'])

    # process output with an API in the subprocess module:
    reqs = check_output([sys.executable, '-m', 'pip','freeze'])
    installed_packages = [r.decode().split('==')[0] for r in reqs.split()]

    print(installed_packages)

try:
    import openpyxl
except ModuleNotFoundError:
    print("\u001b[31m Openpyxl library is not installed.")
    inst_openpyxl()


import openpyxl


#Confirming, finding, and copying files.
def manages_files():
    print("\n")
    wrk_dir = os.getcwd()
    temp_dir = wrk_dir + '//template'
    out_dir = wrk_dir + '//output'
    in_dir = wrk_dir + '//input'
    #confirming files
    try:
        temp_loc = temp_dir + '//' + os.listdir(temp_dir)[0]
    except FileNotFoundError:
        print("\u001b[1m\u001b[31mThe template file or directory was not found.\n\nPlease add the template file to the template directory and restart.")
        done()
    try:
        in_loc = in_dir + '//' + os.listdir(in_dir)[0]
    except FileNotFoundError:
        print("\u001b[1m\u001b[31mThe input file or directory was not found.\n\nPlease add the input file to the input directory and restart.")
        done()
    #Copying template file to output directory
    try:
        shutil.copy(temp_loc, out_dir + '//out_' + os.listdir(temp_dir)[0])
    except FileNotFoundError:
        print("\u001b[1m\u001b[31mThe output directory was not found.\n\nPlease add the output directory and restart.")
        done()
    out_loc = out_dir + '//' + os.listdir(out_dir)[0]
    locations = [temp_loc, out_loc, in_loc]
    return locations


def main():
    file_locs = manages_files()#file_locs[template, output, input]
    
    wb = openpyxl.load_workbook(filename=file_locs[1])
    ws = wb["Import Cheat Sheet"]

    address_array = []
    for i in range(ws.max_row):
        addy = ws.cell(i+3,2).value
        if type(addy) != str:
            continue
        elif addy[0:3] == "GMF":
            address_array.append([addy, i + 3])# [0]requested address [1]position
        elif addy[0:2] == "EM":
            address_array.append([addy, i + 3])
        else:
            pass
    """
    print(address_array)
    print(len(address_array))
    """
    #open and read the csv file.
    try:
        comment_array = list(csv.reader(open(file_locs[2], encoding= "ISO8859")))
    except PermissionError:
        print("\u001b[1m\u001b[31mError: Could not access input file.")
    #print(len(comment_array))
    match_count = 0
    array_len = len(address_array)
    print("\u001b[0m\u001b[36mProgress...")
    #loop through the addresses and compare to the csv list
    for i in range(array_len):
        for comment in comment_array:
            if address_array[i][0] == comment[0]:
                ws.cell(row=address_array[i][1], column=6).value = comment[0]
                ws.cell(row=address_array[i][1], column=7).value = comment[1]
                match_count+=1
            elif i % 525 == 0:#update progress
                #sys.stdout.write(u"\u001b[0m\u001b[37m\u001b[1000D" + str('{:.2f}'.format((i / array_len) * 100)) + "%")
                #sys.stdout.flush()
                continue
            else:
                pass
    #sys.stdout.write(u"\u001b[0m\u001b[37m\u001b[1000D" + str(100.00) + "%")
    #sys.stdout.flush()
    print("\n\u001b[0m\u001b[37mNumber of comments wrote:\u001b[32m" + str(match_count))
    if match_count == 0:
        print("\u001b[33m***No matches were found. Make sure your input and template files are correct***")
    wb.save(file_locs[1])

    done()

#Resets the text color
def done():
    print("\u001b[37m\u001b[0m")
    exit()


main()