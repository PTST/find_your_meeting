import numpy as np
import json
import sys
from pathfinding.core.diagonal_movement import DiagonalMovement
from pathfinding.core.grid import Grid
from pathfinding.finder.a_star import AStarFinder
import openpyxl
import os
import time
import concurrent.futures
import find_room

'''
Import variables from external files
'''
operating_system = sys.platform
def import_vars():
    floor_plans_dict = {}
    if operating_system == "darwin":
        floor_plans_dict[0] = list(np.load("resources/Basement.npy"))
        floor_plans_dict[1] = list(np.load("resources/GroundFloor.npy"))
        floor_plans_dict[2] = list(np.load("resources/SecondFloor.npy"))
        floor_plans_dict[3] = list(np.load("resources/Tower.npy"))
        with open("resources/locations.json", "r") as f:
            jsonfile = f.read()
        locations_list = json.loads(jsonfile)
        with open("resources/up_down.json", "r") as f:
            jsonfile = f.read()
        up_down_list = json.loads(jsonfile)
    elif operating_system == "win32":
        floor_plans_dict[0] = list(np.load("M:\\GitHub\\find_your_meeting\\resources\\Basement.npy"))
        floor_plans_dict[1] = list(np.load("M:\\GitHub\\find_your_meeting\\resources\\GroundFloor.npy"))
        floor_plans_dict[2] = list(np.load("M:\\GitHub\\find_your_meeting\\resources\\SecondFloor.npy"))
        floor_plans_dict[3] = list(np.load("M:\\GitHub\\find_your_meeting\\resources\\Tower.npy"))
        with open("M:\\GitHub\\find_your_meeting\\resources\\locations.json", "r") as f:
            jsonfile = f.read()
        locations_list = json.loads(jsonfile)
        with open("M:\\GitHub\\find_your_meeting\\resources\\up_down.json", "r") as f:
            jsonfile = f.read()
        up_down_list = json.loads(jsonfile)
    elif operating_system == "linux":
        floor_plans_dict[0] = list(np.load("/pytest/resources/Basement.npy"))
        floor_plans_dict[1] = list(np.load("/pytest/resources/GroundFloor.npy"))
        floor_plans_dict[2] = list(np.load("/pytest/resources/SecondFloor.npy"))
        floor_plans_dict[3] = list(np.load("/pytest/resources/Tower.npy"))
        with open("/pytest/resources/locations.json", "r") as f:
            jsonfile = f.read()
        locations_list = json.loads(jsonfile)
        with open("/pytest/resources/up_down.json", "r") as f:
            jsonfile = f.read()
        up_down_list = json.loads(jsonfile)
    return floor_plans_dict, locations_list, up_down_list

def select_start_end(loc_dict):
    while True:
        while True:
            try:
                start_point = input("\nStart: ").upper()
                test = loc_dict[start_point][0]
                break
            except KeyError:
                print("\nNot a valid starting point. Try again")

        while True:
            try:
                end_point = input("\nEnd: ").upper()
                test2 = loc_dict[end_point][0]
                break
            except KeyError:
                print("\nNot a valid endpoint. Try again")
        if test == test2:
            print("Starting point and endpoint is the same")
        else:
            break
    return start_point, end_point

def load_excel():
    if operating_system == "win32":
        wb = openpyxl.load_workbook("M:\\GitHub\\find_your_meeting\\floorPlan.xlsm")
    elif operating_system == "darwin":
        wb = openpyxl.load_workbook("/Users/PTST/Dev/find_your_meeting/floorPlan.xlsm")
    else:
        raise ValueError("Unknown operating system")
    return wb

def save_to_excel(wb, path, floors):
    for i in range(len(path)):
        sheet = wb.get_sheet_by_name(wb.get_sheet_names()[int(floors[i])])
        for item in path[i]:
            sheet.cell(row=(item[1]+1), column=(item[0]+1)).value = "x"
            sheet.cell(row=(item[1]+1), column=(item[0]+1)).font = openpyxl.styles.Font(color=openpyxl.styles.colors.RED)
            sheet.cell(row=(item[1]+1), column=(item[0]+1)).fill = openpyxl.styles.PatternFill(fgColor=openpyxl.styles.colors.RED, fill_type = "solid")
    if operating_system == "win32":
        wb.save("M:\\GitHub\\find_your_meeting\path.xlsx")
        os.startfile("M:\\GitHub\\find_your_meeting\path.xlsx")
    elif operating_system == "darwin":
        wb.save("/Users/PTST/Dev/find_your_meeting/path.xlsx")
        os.system("open /Users/PTST/Dev/find_your_meeting/path.xlsx")
    else:
        raise ValueError("Unknown operating system")

if __name__ == '__main__':

    floor_plans, locations, up_down = import_vars()
    start_loc, end_loc = select_start_end(locations)
    best_path = find_room.run(start_loc, end_loc, floor_plans, locations, up_down)
    workbook = load_excel()
    save_to_excel(workbook, best_path, [start_loc[1:2], end_loc[1:2]])