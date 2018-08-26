import xlrd
import os
import sys
import re
import csv
from datetime import datetime
import xlsxwriter

EXCEL_FILE = "D0920ZU.xls"
FOLDER_NAME = "D0920ZU"
FIELD_NAMES_FILE = "D0920ZU_field_names.csv"
OUTPUT_FOLDER = "opendata"
NOT_SPACE = re.compile("\S+")
GARBAGE = ["Україна", "Міненерговугілля України", "Підприємства МЕВ", "Підприємства МEВ"]
REGION_CODES = {'Львiвська область': "4600000000", "Донецька область": "1400000000", "Луганська область": "4400000000", "Днiпропетровська область": "1200000000", "Волинська область": "0700000000"}
MONTHES_DICT = {"01":"січень" , "02":"лютий", "03":"березень", "04":"квітень", "05":"травень", "06":"червень", "07":"липень", "08":"серпень", "09":"вересень", "10":"жовтень", "11":"листопад", "12":"грудень"}
COMPANY_CODES =  {'ДП ш/у Пiвденнодонбаське1':"34032208", 'ПАТ"Шахтоуправління Покровське"':"13498562", 'ДП "Шахта ім.М.С.Сургая "':"40695853", 'ДП Мирноградвугілля':"32087941", 'ДП "ВК Краснолиманська"':"31599557", 'ДП Селидiввугiлля': "33426253", 'ТДВ " ш Бiлозерська"': "36028628", 'ТОВ Краснолиманська': "36144814", 'ДП Первомайськвугiлля':"32320594", 'ПАТ ш.Надiя': "00178175", 'ПАТ "ДТЕК Павлоградвугiлля"': "00178353", 'ДП Львiввугiлля': "32323256", 'ТОВ "ДТЕК Добропiллявугiлля"':"37014600", 'ДП Торецьквугiлля':"33839013", 'ДП Волиньвугiлля':"32365965", 'ПАТ "Лисичанськвугiлля"':"32359108", 'ТОВ "ДТЕК  Комсомолець Донбасу"': "05508186", 'ТДВ"ОП"Ш. ім Святої Матрони Московскої':"36182252", 'ТОВ "ДТЕК Ровенькиантрацит"':'37713861', 'ТОВ "ДТЕК Свердловантрацит"':"37596090", 'ДП Красноармiйськвугiлля': "32087941"}
ADD_NAMES = ["32359108", "00178175"]
HEADERS = ["month", "company", "company_code", "mine", "ministry_owned_company", "region", "region_code", "mark", "extraction_fact", "extraction_plan", "ash_percent_fact", "ash_percent_plan"]
FILENAME_TEMPLATE = "coal_extraction_{month:s}"


def is_row_initial(row):
    global ncol, sheet
    values = []
    for i in range(ncol):
        values.append(sheet.cell(row, i).value)
    return values[:4] == [1, 2, 3, 4]

def month_row(row_number):
    global MONTHES_DICT
    cellvalue = sheet.cell(row_number, 0).value
    m = [MONTHES_DICT[k] in str(cellvalue) for k in MONTHES_DICT.keys()]
    return sum(m) > 0

def is_italic(cell):
    font_index = wb.xf_list[cell.xf_index].font_index
    return wb.font_list[font_index].italic == 1

def is_size_equals(cell, size):
    font_index = wb.xf_list[cell.xf_index].font_index
    return wb.font_list[font_index].height == size

def is_blank(cell):
    return NOT_SPACE.search(str(cell.value)) == None

def is_enterprise(cell):
    return (not is_blank(cell)) and is_italic(cell) and is_size_equals(cell, 200)

def is_region(cell):
    return (not is_blank(cell)) and "обл." in cell.value

def dump_row(row_number):
    global sheet_dict, region, dp, month
    for k in fields_dictionary.keys():
        sheet_dict[fields_dictionary[k]].append(sheet.cell(row_number, int(k) - 1).value)
    sheet_dict['region'].append(region)
    sheet_dict['company'].append(dp)
    sheet_dict['month'].append(month)
    

def load_workbook(wb):
    global sheet, ncol, dp, region, month
    sheet = wb.sheet_by_index(0)
    ncol = sheet.ncols
    nrows = sheet.nrows
    i = 0
    data_rows = False
    while i < nrows:
        ind = sheet.cell_xf_index(i,0)
        if not data_rows:
            if month_row(i):
                cellvalue = sheet.cell(i, 0).value
                parts = cellvalue.split()
                month = monthes_dict[parts[1]] + "." + parts[2]
            elif is_row_initial(i):
                data_rows = True
        else:
            if not(sheet.cell(i, 0).value in GARBAGE):
                #print(is_enterprise(sheet.cell(i, 0)), sheet.cell(i, 0).value)
                if is_region(sheet.cell(i, 0)):
                    region = sheet.cell(i, 0).value.replace("обл.", "область")
                elif is_enterprise(sheet.cell(i, 0)):
                    dp = sheet.cell(i, 0).value
                    if is_blank(sheet.cell( i + 1, 0)):
                        dump_row(i)
                elif not is_blank(sheet.cell(i, 0)):
                    dump_row(i)
        i += 1

def dict_to_list(dict_, headers):
    l = []
    for i in range(len(dict_[headers[0]])):
        new_l = []
        for h in headers:
            new_l.append(dict_[h][i])
        l.append(new_l)
    return l

if not os.path.exists("coal_extraction.log"):
    f = open("coal_extraction", "w")
    f.close()
f = open("coal_extraction", "a")
sys.stdout = f
sys.stderr = f

print("----------------")
print(datetime.now())

with open(FIELD_NAMES_FILE, 'r') as vf:
    var_reader = csv.reader(vf)
    fields_dictionary = {}
    for l in var_reader:
        fields_dictionary[l[0]] = l[2]
sheet_dict = {}
for k in fields_dictionary.keys():
    sheet_dict[fields_dictionary[k]] = []
sheet_dict['region'] = []
sheet_dict['company'] = []
sheet_dict['month'] = []

monthes_dict = {}
for k in MONTHES_DICT:
    monthes_dict[MONTHES_DICT[k]] = k

excel_files = os.listdir(FOLDER_NAME)
excel_files = [f for f in excel_files if f.endswith(".xls")]
for f in excel_files:
    print(f)
    wb = xlrd.open_workbook(os.path.join(FOLDER_NAME, f), formatting_info=True)
    load_workbook(wb)
sheet_dict['company_code'] = list(map(lambda c: COMPANY_CODES[c], sheet_dict['company']))
sheet_dict['month'] = [datetime.strptime(m, "%m.%Y") for m in sheet_dict['month']]
month_to_filename = max(sheet_dict['month']).strftime("%m_%Y")
filename = os.path.join(OUTPUT_FOLDER, FILENAME_TEMPLATE.format(month = month_to_filename))
sheet_dict['ministry_owned_company'] = [1 for i in range(len(sheet_dict['month']))]
sheet_dict['region_code'] = list(map(lambda r: REGION_CODES[r], sheet_dict['region']))
indices = []
for i in range(len(sheet_dict['company'])):
    if "ДП" in sheet_dict['company'][i] or sheet_dict['company_code'][i] in ADD_NAMES:
        indices.append(i)
coal_list = dict_to_list(sheet_dict, HEADERS)
coal_list = [coal_list[i] for i in indices]
coal_list = sorted(coal_list, key=lambda x: x[0],reverse=True)
with open(filename + ".csv", "w", newline="") as cfile:
    csvwriter = csv.writer(cfile)
    csvwriter.writerow(HEADERS)
    for i in range(len(coal_list)):
        l = coal_list[i][:]
        l[0] = datetime.strftime(l[0],"%m.%Y")
        csvwriter.writerow(l)

out_wb = xlsxwriter.Workbook(filename + ".xlsx")
worksheet = out_wb.add_worksheet()
datef = out_wb.add_format({'num_format':"mm yyyy"})
ashf = out_wb.add_format({'num_format':"0.00"})
headerf = out_wb.add_format({'bold':True})
for i in range(len(HEADERS)):
    worksheet.write(0, i, HEADERS[i], headerf)
for i in range(len(coal_list)):
    for j in range(len(HEADERS)):
        if j == 0:
            worksheet.write(i+1, j, coal_list[i][j], datef)
        elif j >  9:
            worksheet.write(i+1, j, coal_list[i][j], ashf)
        else:
            worksheet.write(i+1, j, coal_list[i][j])
out_wb.close()

print("No errors were caught")
print("-----------------")