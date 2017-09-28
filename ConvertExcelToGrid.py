import openpyxl


'''
OQ
'''

wb = openpyxl.load_workbook('/Users/PTST/Dev/find_your_meeting/groundFloor.xlsx')
sheet = wb.get_sheet_by_name(wb.get_sheet_names()[0])
maxCell = openpyxl.utils.column_index_from_string("OQ") 

master_list = []
location_dict = {}

for y in range(1, maxCell+1):
    x_list = []
    for x in range(1, maxCell+1):
        if sheet.cell(row=y, column=x).value == "#":
            x_list.append(1)
        elif sheet.cell(row=y, column=x).value != "":
            x_list.append(0)
            location_dict[sheet.cell(row=y, column=x).value] = [x, y]
        else:
            x_list.append(0)
    master_list.append(x_list)

for key, value in location_dict.items() :
    print (key, value)

