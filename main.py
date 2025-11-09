from PIL import Image
import openpyxl
from openpyxl.styles import fills

work_book = openpyxl.Workbook()
work_sheet = work_book.active


img = Image.open("images.jpg")
pix = img.load()


def cell_finder(y, x):
    for y_cell in range(1, y):
        for x_cell in range(1, x):
            cell_name = work_sheet.cell(row=y_cell, column=x_cell)
            cell_coordinate = cell_name.coordinate
            return cell_coordinate
# for image loop
for y in range(0, 255, 10):
    for x in range(0, 198, 10):
        pixels = pix[y, x]
        r, g, b = pixels
        hex_color = "{:02X}{:02X}{:02X}".format(r, g, b)
        
        fill = fills.PatternFill(
            start_color="ff" + hex_color, end_color="ff" + hex_color, fill_type="solid"
        )
        

        # for excel loop
        for y_cell in range(1, y):
            for x_cell in range(1, x):
                cell_name = work_sheet.cell(row=y_cell, column=x_cell)
                cell_coordinate = cell_name.coordinate
                
        work_sheet[f"{cell_coordinate}"].fill = fill

work_book.save("image_excel.xlsx")
