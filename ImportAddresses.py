from codecs import raw_unicode_escape_decode
import os
import csv

print("Fault Address Layout Import")

def inst_openpyxl():
    None

try:
    import openpyxl
except ModuleNotFoundError:
    print("Openpyxl library is not installed.")
    inst_openpyxl()


def main():
    check_perm()

    wb = openpyxl.load_workbook(filename="BW Specific Fault Layout Toyopuc V9.xlsx")
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


    comment_array = list(csv.reader(open("TMMI_UB_Respot_Main_20220827.csv", encoding= "ISO8859")))
    #print(len(comment_array))
    match_count = 0
    for i in range(len(address_array)):
        for comment in comment_array:
            if address_array[i][0] == comment[0]:
                ws.cell(row=address_array[i][1], column=6).value = comment[0]
                ws.cell(row=address_array[i][1], column=7).value = comment[1]
                match_count+=1
            else:
                continue
    print(match_count)
    wb.save("BW Specific Fault Layout Toyopuc V9.xlsx")


def check_perm():
    print("Going to check a few things before we start...")
    print("BW Specific Fault Layout Toyopuc V9.xlsx... file exsists =", os.access("BW Specific Fault Layout Toyopuc V9.xlsx", os.F_OK))
    print("BW Specific Fault Layout Toyopuc V9.xlsx... read access =", os.access("BW Specific Fault Layout Toyopuc V9.xlsx", os.R_OK))
    print("BW Specific Fault Layout Toyopuc V9.xlsx... write access =", os.access("BW Specific Fault Layout Toyopuc V9.xlsx", os.W_OK))
    print("TMMI_UB_Respot_Main_20220827.csv... file exsists =", os.access("TMMI_UB_Respot_Main_20220827.csv", os.F_OK))
    print("TMMI_UB_Respot_Main_20220827.csv... read access =", os.access("TMMI_UB_Respot_Main_20220827.csv", os.F_OK))


main()