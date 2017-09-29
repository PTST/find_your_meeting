import openpyxl
from pathfinding.core.diagonal_movement import DiagonalMovement
from pathfinding.core.grid import Grid
from pathfinding.finder.a_star import AStarFinder
import sys
import os
import copy

"""
OQ
"""
def load_excel(lastColumn):
    operating_system = sys.platform
    if operating_system == "win32":
        wb = openpyxl.load_workbook("M:\\GitHub\\find_your_meeting\\groundFloor.xlsm")
    elif operating_system == "darwin":
        wb = openpyxl.load_workbook("/Users/PTST/Dev/find_your_meeting/groundFloor.xlsm")
    else:
        raise ValueError("Unknown operating system")
    return wb, wb.get_sheet_by_name(wb.get_sheet_names()[0]), openpyxl.utils.column_index_from_string(lastColumn)

def convert_to_grid(sheet, last_cell):
    master_list = []
    location_dict = {}
    for y in range(1, last_cell+1):
        x_list = []
        for x in range(1, last_cell+1):
            if sheet.cell(row=y, column=x).value == ".":
                x_list.append(0)
            elif sheet.cell(row=y, column=x).value is not None:
                location_dict[str(sheet.cell(row=y, column=x).value).lower()] = [x-1, y-1]
                x_list.append(0)
            else:
                x_list.append(1)
        master_list.append(x_list)
    grid = Grid(matrix=master_list)
    return grid, location_dict

def print_locations(location_dict):
    print("\nPossible locations:")
    for key, value in location_dict.items() :
        print (key, value)


def find_path(location_dict, grid):
    while True:
        while True:
            try:
                start_node = input("\nStart: ").lower()
                start = grid.node(location_dict[start_node][0], location_dict[start_node][1])
                break
            except KeyError:
                print("\nNot a valid starting point. Try again")
                print_locations(location_dict)

        while True:
            try:
                end_node = input("\nEnd: ").lower()
                end = grid.node(location_dict[end_node][0], location_dict[end_node][1])
                break
            except KeyError:
                print("\nNot a valid endpoint. Try again")
                print_locations(location_dict)
        if end == start:
            print("Starting point and endpoint is the same")
        else:
            break

    finder = AStarFinder(diagonal_movement=DiagonalMovement.never)
    path, runs = finder.find_path(start, end, grid)

    print("operations:", runs, "path length:", len(path))
    return path


def save_to_excel(wb, path, sheet):
    operating_system = sys.platform
    for item in path:
        sheet.cell(row=(item[1]+1), column=(item[0]+1)).value = "x"
        sheet.cell(row=(item[1]+1), column=(item[0]+1)).font = openpyxl.styles.Font(color=openpyxl.styles.colors.RED)
        sheet.cell(row=(item[1]+1), column=(item[0]+1)).fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.RED, fill_type = "solid")
    if operating_system == "win32":
        wb.save("M:\\GitHub\\find_your_meeting\path.xlsx")
    elif operating_system == "darwin":
        wb.save("/Users/PTST/Dev/find_your_meeting/path.xlsx")
    else:
        raise ValueError("Unknown operating system")

    os.startfile("M:\\GitHub\\find_your_meeting\path.xlsx")


last_column=input("Last column: ").upper()
workbook, excel_sheet, max_ccell = load_excel(last_column)
orig_workbook = copy.copy(workbook)
search_grid, possible_locations = convert_to_grid(excel_sheet, max_ccell)
do = True
while do:
    print_locations(possible_locations)
    found_path = find_path(possible_locations, search_grid)
    save_to_excel(workbook, found_path, excel_sheet)
    while True:
        answer = input("Run again (Y/N): ").lower()
        if (answer == "y") or (answer == "n"):
            do = (answer == "y")
            workbook, excel_sheet, max_ccell = load_excel(last_column)
            break

