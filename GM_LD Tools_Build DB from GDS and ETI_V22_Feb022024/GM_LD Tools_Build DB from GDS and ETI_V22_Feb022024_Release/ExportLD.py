from bs4 import BeautifulSoup
import os
import pandas as pd

#######################
# Function xử lý:


def GetLDgp_Name(data):
    list_LDgp_Name = []
    for i in data.find_all('h2'):
        list_LDgp_Name.append(i.string)
    LDGrpname = list_LDgp_Name[3]
    return LDGrpname


def READ_HTML(data, LD_Group_Name):
    # nơi chứa data của 1 file html:
    DF_1_HTML = pd.DataFrame({"Group": ['x'], "PID Name": ['x'], "Unit": ['x'], "Key Check GDS Sp": ['x']})

    # print(data)
    list_table = []
    for i in data.find_all('table'):
        list_table.append(i)
    LD_DATA = list_table[3]
    # print(LD_DATA)
    List_Tab_Tr = []
    for y in LD_DATA.find_all('tr'):
        List_Tab_Tr.append(y)

    for xx in List_Tab_Tr:
        # print("tab tr: ", xx)
        List_Tab_Td = []
        for aa in xx.find_all('td'):
            List_Tab_Td.append(aa.string)
        # print(List_Tab_Td)

        if len(List_Tab_Td) > 0:
            keycheck = str(List_Tab_Td[1]) + str(List_Tab_Td[3])
            DF_1_Tab_Td = pd.DataFrame({"Group": [LD_Group_Name], "PID Name": [List_Tab_Td[1]], "Unit": [List_Tab_Td[3]], "Key Check GDS Sp": [keycheck]})
            frame = [DF_1_HTML, DF_1_Tab_Td]
            DF_1_HTML = pd.concat(frame)
    # print("=====>\n", DF_1_HTML)
    return(DF_1_HTML)


#######################


def run_ExportLD_Group(PATHX):
    DF_YMME_HTML = pd.DataFrame({"Group": ['x'], "PID Name": ['x'], "Unit": ['x'], "Key Check GDS Sp": ['x']})
    YMME = PATHX.replace('DATA/', '')
    for filename in os.listdir(PATHX):
        # print("--->", filename)
        if filename.endswith(".html"):
            Path_LDgrp = PATHX + '/' + str(filename)
            # print("----> Path: ", Path_LDgrp)
            try:
                html_file = BeautifulSoup(open(Path_LDgrp), features="html.parser")
                LD_Group_Name = GetLDgp_Name(html_file)
                print("LD_Group_Name :", LD_Group_Name)
                DATA_1_Group = READ_HTML(html_file, LD_Group_Name)
                DF_YMME_HTML = pd.concat([DF_YMME_HTML, DATA_1_Group])
            except:
                pass

    data = DF_YMME_HTML.drop_duplicates(subset=None, keep='first')
    # print("Your Data in this YMME: \n", data)
    # file_path_data = PATHX + '/' + str(YMME) + 'LDgroup' + '.xlsx'
    # data.to_excel(file_path_data, index=False)
    return data


# ====
# PATHX = 'DATA/2021GMBuickEnclave'
# run_ExportLD_Group(PATHX)
