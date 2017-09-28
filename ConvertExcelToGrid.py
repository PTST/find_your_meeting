import openpyxl


'''
OQ
'''

wb = openpyxl.load_workbook('/Users/PTST/Dev/find_your_meeting/groundFloor.xlsx')
sheet = wb.get_sheet_by_name(wb.get_sheet_names()[0])
maxCell = openpyxl.utils.column_index_from_string("OQ")
print(maxCell)