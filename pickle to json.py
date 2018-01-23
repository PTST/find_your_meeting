import numpy as np
import json
import codecs
floor_plans_dict = {}
floor_plans_dict[0] = list(np.load("M:\\GitHub\\find_your_meeting\\resources\\Basement.npy").tolist())
floor_plans_dict[1] = list(np.load("M:\\GitHub\\find_your_meeting\\resources\\GroundFloor.npy").tolist())
floor_plans_dict[2] = list(np.load("M:\\GitHub\\find_your_meeting\\resources\\SecondFloor.npy").tolist())
floor_plans_dict[3] = list(np.load("M:\\GitHub\\find_your_meeting\\resources\\Tower.npy").tolist())

file_path = "M:\path.json"  # your path variable
with open('M:\data.txt', 'w') as outfile:
    json.dump(floor_plans_dict, outfile)

