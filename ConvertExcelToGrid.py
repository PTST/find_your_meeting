import openpyxl
from pathfinding.core.diagonal_movement import DiagonalMovement
from pathfinding.core.grid import Grid
from pathfinding.finder.a_star import AStarFinder

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

grid = Grid(matrix=master_list)

startNode = "h1.25d"
endNode = "kantinen"

location_dict[startNode][0], location_dict[startNode][1]

start = grid.node(location_dict[startNode][0], location_dict[startNode][1])
end = grid.node(location_dict[endNode][0], location_dict[endNode][1])

finder = AStarFinder(diagonal_movement=DiagonalMovement.always)
path, runs = finder.find_path(start, end, grid)

print('operations:', runs, 'path length:', len(path))
for item in path:
    sheet.cell(row=(item[1]+1), column=(item[0]+1)).value = "x"

wb.save('/Users/PTST/Dev/find_your_meeting/path.xlsx')

