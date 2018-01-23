import numpy as np
import json
import sys
from pathfinding.core.diagonal_movement import DiagonalMovement
from pathfinding.core.grid import Grid
from pathfinding.finder.a_star import AStarFinder
import os
import time

operating_system = sys.platform
def find_path(location_dict, start_point, end_point, grid):
    start = grid.node(location_dict[start_point][0], location_dict[start_point][1])
    end = grid.node(location_dict[end_point][0], location_dict[end_point][1])
    finder = AStarFinder(diagonal_movement=DiagonalMovement.never)
    path, runs = finder.find_path(start, end, grid)
    return path

def multiple_floors(counter, floor_no, going_up, plan_dict, paths_start_dict, paths_end_dict, up_down_locations, possible_locations, starting_loc, ending_loc):
    floor = plan_dict[floor_no]
    
    if counter == 0:
        if going_up:
            direction = "UP"
        else:
            direction = "DOWN"
    else:
        if going_up:
            direction = "DOWN"
        else:
            direction = "UP"

    for key, value in up_down_locations.items():
        get_floor_paths(key, direction, floor_no, floor, counter, paths_start_dict, paths_end_dict, possible_locations, starting_loc, ending_loc)

    return paths_start_dict, paths_end_dict


def get_floor_paths(a, up_or_down, floor_idx, floor_plan, count, found_paths_start, found_paths_end, locations_list, start_location, end_location):
    if a.endswith(up_or_down) and a.startswith(str(floor_idx)):
        search_grid = Grid(matrix=floor_plan)
        if count == 0:
            found_path = find_path(locations_list, start_location, a, search_grid)
            if len(found_path) != 0:
                found_paths_start[a] = found_path
        elif count == 1:
            found_path = find_path(locations_list, a, end_location, search_grid)
            if len(found_path) != 0:
                found_paths_end[a] = found_path
        
    return found_paths_start, found_paths_end

def run(start_loc, end_loc, floor_plans, locations, up_down):
    start_time = time.time()
    start_floor, end_floor = start_loc[1], end_loc[1]

    if start_floor != end_floor:
        paths_start = {}
        paths_end = {}
        upwards = (start_floor < end_floor)
        for idx, floor_id in enumerate([int(start_floor), int(end_floor)]):
            paths_start, paths_end = multiple_floors(idx, floor_id, upwards, floor_plans, paths_start, paths_end, up_down, locations, start_loc, end_loc)
            

    best_value = sys.maxsize
    best_path = []
    for key, value in paths_start.items():
        for key2, value2 in paths_end.items():
            if key[2:3] == key2[2:3]:
                total = len(value) + len(value2)
                if total < best_value:
                    best_value = total
                    best_path = []
                    best_path.append(value)
                    best_path.append(value2)

    print("--- %s seconds ---" % (time.time() - start_time))
    return(best_path)

if __name__ == "__main__":
    pass