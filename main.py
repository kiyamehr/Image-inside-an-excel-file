from PIL import Image
import openpyxl
from openpyxl.styles import PatternFill

work_book = openpyxl.Workbook()
work_sheet = work_book.active

img = Image.open("images.jpg").convert("RGB")
pix = img.load()


# for image loop
for y in range(0, 198, 1):
    for x in range(0, 255, 1):
        r, g, b = pix[x, y]

        hex_color = "{:02X}{:02X}{:02X}".format(r, g, b)
        fill = PatternFill(
            start_color="FF" + hex_color,
            end_color="FF" + hex_color,
            fill_type="solid"
        )

        # put directly instead of finding separately
        work_sheet.cell(row=y // 10 + 1, column=x // 10 + 1).fill = fill




work_book.save("image_excel.xlsx")
