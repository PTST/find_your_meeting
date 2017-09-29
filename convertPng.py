from PIL import Image

import openpyxl
wb = openpyxl.load_workbook('/Users/PTST/Dev/find_your_meeting/test.xlsx')
sheet = wb.get_sheet_by_name(wb.get_sheet_names()[0])

im = Image.open("/Users/PTST/Dev/find_your_meeting/basement.png") #Can be many different formats.
pix = im.load()
img_x, img_y = tuple(map(int, im.size))

white = (255,255,255,255)
print(pix[5,5] != white)
for y in range(0, img_y, 3):
    for x in range(0, img_x, 3):
        if pix[x,y] != white:
            sheet.cell(row=int((y+1)/3), column=int((x+1)/3)).fill = openpyxl.styles.PatternFill(fgColor="000000", fill_type = "solid")
wb.save('/Users/PTST/Dev/find_your_meeting/basement.xlsx')
