import openpyxl
import copy
import numpy as np
import json

wb = openpyxl.load_workbook("Basement2.xlsx")

basement, ground_floor, first_floor, tower = [], [], [], []

column_dict = {"Basement":openpyxl.utils.column_index_from_string("AUU"), "GroundFloor":407, "SecondFloor":344, "Tower":70}

location_dict = {}
up_down_dict = {}

for sheet_name in wb.get_sheet_names():
    last_cell = column_dict[sheet_name]
    sheet = wb.get_sheet_by_name(sheet_name)
    master_list = []

    for y in range(1, last_cell+1):
        x_list = []
        for x in range(1, last_cell+1):
            cell_value = sheet.cell(row=y, column=x).value
            if cell_value == ".":
                x_list.append(0)
            elif cell_value is not None:
                if (cell_value.upper().endswith("UP")) or (cell_value.upper().endswith("DOWN")):
                    up_down_dict[str(cell_value).upper()] = [x-1, y-1]
                    location_dict[str(cell_value).upper()] = [x-1, y-1]
                else:
                    location_dict[str(cell_value).upper()] = [x-1, y-1]
                x_list.append(0)
            else:
                x_list.append(1)
        master_list.append(x_list)
    
    np.save("resources/"+sheet_name, master_list)
    print(sheet_name + " - Done")

jsonFile = json.dumps(location_dict) # note i gave it a different name
with open("resources/locations.json", "w") as f:
    f.write(jsonFile)
    print("locations.json - done")
jsonFile = json.dumps(up_down_dict) # note i gave it a different name
with open("resources/up_down.json", "w") as f:
    f.write(jsonFile)
    print("up_down.json - done")