import xlrd
import os
import re
import csv
from datetime import datetime
import xlsxwriter
import sys

FOLDER_NAME = "doc04"
OUTPUT_FOLDER = "opendata"
EXCEL_FILE = "doc04.xls"
FIELD_NAMES_FILE = "doc04_field_names.csv"
FILENAME_TEMPLATE = "coal_refining_{date:s}"
NOT_SPACE = re.compile("\S+")
COMPANY_CODES =  {'ДП"ш/у Пiвденнодонбас':"34032208", 'ДП"Волиньвугiлля"':"32365965", 'ДП"Мирноградвугiлля"':"32087941", 'ДП "Первомайськвугiлл':"32320594", 'ДП"Селидiввугiлля"': "33426253", 'ПАТ "Лисичанськвугiлл':"32359108", 'ГП "Львiввугiлля"': "32323256", 'ДП "Торецьквугiлля"':"33839013", 'ДП "Добропiллявугiлля':"37014600", 'ТОВ "ДТЕК Добропiллявугiлля"':"37014600", 'ТОВ "ДТЕК СА"': "37596090", 'ДП "Львiввугiлля"':"32323256", 'ГП "Львiввугiлля"':"32323256", 'ДП"Красноармiйськвугi': "32087941", 'ОДО "ш.Белозерская"': "36028628", 'ТДВ ш. "Білозерська"': "36028628"}
MONTH_DICTIONARY = {"січня":"01", "лютого":"02", "березня":"03", "квітня":"04", "травня":"05", "червня":"06", "липня":"07", "серпня":"08", "вересня":"09", "сентября":"09", "жовтня":"10", "листопада":"11", "грудня":"12", "марта": "03", "сiчня": "01"}
HEADERS = ["date", "company", "company_code", "ministry_owned_company", "mine", "mark", "coal_sent_month_plan", "coal_sent_month_fact", "coal_sent_year_plan", "coal_sent_year_fact"]
HEADERS_SHORT = ["date", "company", "mine", "mark", "coal_sent_month_plan", "coal_sent_month_fact", "coal_sent_year_plan", "coal_sent_year_fact"]
DATE_RE = re.compile("\d{2}\s+[a-zа-яіїє]+\s+\d{4}")
ADD_NAMES = ["32359108", "00178175", "36028628", "32323256", "37014600"]

def is_company(cell):
    ind = cell.xf_index
    return wb.xf_list[ind].background.pattern_colour_index == 40

def is_italic(cell):
    font_index = wb.xf_list[cell.xf_index].font_index
    return wb.font_list[font_index].italic == 1

def is_blank(cell):
    return NOT_SPACE.search(str(cell.value)) == None

def dict_to_list(dict_, headers):
    l = []
    for i in range(len(dict_[headers[0]])):
        new_l = []
        for h in headers:
            new_l.append(dict_[h][i])
        l.append(new_l)
    return l

def dump_row(row_number):
    global sheet_dict, dp, date_
    for k in fields_dictionary.keys():
        new_value = sheet.cell(row_number, int(k) - 1).value
        if isinstance(new_value, str):
            sheet_dict[fields_dictionary[k]].append(new_value.strip())
        else:
            sheet_dict[fields_dictionary[k]].append(new_value)
    sheet_dict['company'].append(dp)
    sheet_dict['date'].append(date_)

def summarise_company():
    global sheet_dict, sheet_dict_big
    headers = sheet_dict.keys()
    sum_by = HEADERS_SHORT[:-4:1]
    sum_fields = HEADERS_SHORT[-4::1]
    big_update = {}
    for h in HEADERS_SHORT:
        big_update[h] = []
    summary = {}
    groups_list = dict_to_list(sheet_dict, HEADERS_SHORT[:-4:1])
    sum_list = dict_to_list(sheet_dict, HEADERS_SHORT[-4::1])
    for i in range(len(sheet_dict['date'])):
        if i == 0 or not (groups_list[i] in groups_list[:i]):
            for h in HEADERS_SHORT:
                big_update[h].append(sheet_dict[h][i])
        else:
            big_dict_list = dict_to_list(big_update, HEADERS_SHORT[:-4:1])
            ind = big_dict_list.index(groups_list[i])
            for h in sum_fields:
                big_update[h][ind] = transform_numbers(big_update[h][ind]) +  transform_numbers(sheet_dict[h][i])
    for k in sheet_dict_big.keys():
        sheet_dict_big[k] = sheet_dict_big[k] + big_update[k]
        sheet_dict[k] = []

def load_workbook(wb):
    global sheet, ncol, dp, date_
    sheet = wb.sheet_by_index(0)
    ncol = sheet.ncols
    nrows = sheet.nrows
    dp = ''
    for i in range(nrows):
        cell = sheet.cell(i,0)
        if date_ == "":
            date_re_matched = DATE_RE.search(cell.value)
            if date_re_matched:
                date_ = date_re_matched.group()
                date_splitted = date_.split()
                date_ = "-".join([date_splitted[2], MONTH_DICTIONARY[date_splitted[1]], date_splitted[0]])
        if not is_blank(cell):
            if cell.value.replace(" ", "").lower != "всего":
                if is_company(cell):
                    summarise_company()
                    dp = cell.value.replace("*","").strip()
                elif not is_italic(cell):
                    if dp != "":
                        dump_row(i)
            else:
                break

def transform_numbers(x):
    if NOT_SPACE.search(str(x)) == None:
        return 0
    else:       
        return float(x)

def replace_nulls(x):
    if isinstance(x, float) and x > 0:
        return x

if not os.path.exists("coal_refine.log"):
    f = open("coal_refine", "w")
    f.close()
f = open("coal_refine", "a")
sys.stdout = f
sys.stderr = f

print("----------------")
print(datetime.now())

        
with open(FIELD_NAMES_FILE, 'r', encoding="utf8") as vf:
    var_reader = csv.reader(vf)
    fields_dictionary = {}
    for l in var_reader:
        fields_dictionary[l[0]] = l[2]
sheet_dict = {}
for k in fields_dictionary.keys():
    sheet_dict[fields_dictionary[k]] = []
sheet_dict['company'] = []
sheet_dict['date'] = []
sheet_dict_big = sheet_dict.copy()
files = os.listdir(FOLDER_NAME)
files = [f for f in files if f.endswith(".xls")]
for f in files:
    date_ = ""
    wb = xlrd.open_workbook(os.path.join(FOLDER_NAME, f), formatting_info=True)
    load_workbook(wb)
sheet_dict = sheet_dict_big.copy()
sheet_dict['company_code'] = list(map(lambda c: COMPANY_CODES[c], sheet_dict['company']))
sheet_dict['date'] = [datetime.strptime(m, '%Y-%m-%d') for m in sheet_dict['date']]
date_to_filename = max(sheet_dict['date']).strftime("%d_%m_%Y")
filename = os.path.join(OUTPUT_FOLDER, FILENAME_TEMPLATE.format(date = date_to_filename))
sheet_dict['ministry_owned_company'] = [1 for i in range(len(sheet_dict['date']))]
sheet_dict["coal_sent_month_plan"] = list(map(replace_nulls, sheet_dict["coal_sent_month_plan"]))
sheet_dict["coal_sent_month_fact"] = list(map(replace_nulls, sheet_dict["coal_sent_month_fact"]))
sheet_dict["coal_sent_year_plan"] = list(map(replace_nulls, sheet_dict["coal_sent_year_plan"]))
sheet_dict["coal_sent_year_fact"] = list(map(replace_nulls, sheet_dict["coal_sent_year_fact"]))
indices = []
for i in range(len(sheet_dict['company'])):
    if "ДП" in sheet_dict['company'][i] or sheet_dict['company_code'][i] in ADD_NAMES:
        indices.append(i)
coal_list = dict_to_list(sheet_dict, HEADERS)
coal_list = [coal_list[i] for i in indices]
coal_list = sorted(coal_list, key=lambda x: x[0],reverse=True)
with open(filename + ".csv", "w", newline="", encoding="utf8") as cfile:
    csvwriter = csv.writer(cfile)
    csvwriter.writerow(HEADERS)
    for i in range(len(coal_list)):
        l = coal_list[i][:]
        l[0] = datetime.strftime(l[0],"%d.%m.%Y")
        csvwriter.writerow(l)
out_wb = xlsxwriter.Workbook(filename + ".xlsx")
worksheet = out_wb.add_worksheet()
datef = out_wb.add_format({'num_format':"dd mm yyyy"})
numf = out_wb.add_format({'num_format':"0.00"})
headerf = out_wb.add_format({'bold':True})
for i in range(len(HEADERS)):
    worksheet.write(0, i, HEADERS[i], headerf)
for i in range(len(coal_list)):
    for j in range(len(HEADERS)):
        if j == 0:
            worksheet.write(i+1, j, coal_list[i][j], datef)
        elif j >  5:
            worksheet.write(i+1, j, coal_list[i][j], numf)
        else:
            worksheet.write(i+1, j, coal_list[i][j])
out_wb.close()      

print("No errors were caught")
print("-----------------")
