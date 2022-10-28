import sys
import subprocess
import os
import csv
import string


def inst_openpyxl():
    None


def inst_pandas():
    None


try:
    import pandas as pd
except ModuleNotFoundError:
    print("Pandas library not installed.")
    inst_pandas()
try:
    import openpyxl
except ModuleNotFoundError:
    print("Openpyxl library is not installed.")
    inst_openpyxl()


def main():
    check_perm()
    temp_sheet = pd.read_excel("BW Specific Fault Layout Toyopuc V9.xlsx", skiprows=1)
    #print(temp_sheet)
    pc_win_data = pd.read_csv("TMMI_UB_Respot_Main_20220827.csv", encoding="ISO8859")
    #print(pc_win_data)
    print(temp_sheet.iat[0,1])
    match_count = 0
    gmf_count = 0
    wb = openpyxl.load_workbook(filename="BW Specific Fault Layout Toyopuc V9.xlsx")
    ws = wb["Import Cheat Sheet"]

    address_list = []
    for i in range(ws.max_row):
        addy = ws.cell(i+3,2).value
        if addy[0:3] == "GMF":
            address_list.append(addy)
    print(address_list)

    for x in range(len(temp_sheet)):
        req_add = temp_sheet.iat[x, 1]
        if req_add[0:3] == "GMF":
            gmf_count = gmf_count + 1
            for y in range(len(pc_win_data)):
                if req_add == pc_win_data.iat[y,0]:
                    match_count = match_count + 1
                    ws.cell(row = x + 3 , column = 6).value = req_add
                    break
                else:
                    continue
        else:
            continue
    wb.save("BW Specific Fault Layout Toyopuc V9.xlsx")
    print(match_count)
    print(gmf_count)


def check_perm():
    print("BW Specific Fault Layout Toyopuc V9.xlsx... file exsists =", os.access("BW Specific Fault Layout Toyopuc V9.xlsx", os.F_OK))
    print("BW Specific Fault Layout Toyopuc V9.xlsx... read access =", os.access("BW Specific Fault Layout Toyopuc V9.xlsx", os.R_OK))
    print("BW Specific Fault Layout Toyopuc V9.xlsx... write access =", os.access("BW Specific Fault Layout Toyopuc V9.xlsx", os.W_OK))
    print("TMMI_UB_Respot_Main_20220827.csv... file exsists =", os.access("TMMI_UB_Respot_Main_20220827.csv", os.F_OK))
    print("TMMI_UB_Respot_Main_20220827.csv... read access =", os.access("TMMI_UB_Respot_Main_20220827.csv", os.F_OK))


main()