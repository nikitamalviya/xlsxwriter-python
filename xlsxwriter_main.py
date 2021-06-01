# from posix import XATTR_REPLACE
import xlsxwriter, os, sys
from glob import glob
from datetime import datetime
# from os.path import normpath

def main(input_path):
    if not os.path.exists("output/"):
        os.mkdir("output/")
        
    xlsx_path = "output\\"+  input_path.split("\\")[-1] + ".xlsx"
    if os.path.exists(xlsx_path):
        os.remove(xlsx_path)
        print("Removing old XLSX file...")
    
    # creating file obj and file
    workbook = xlsxwriter.Workbook(xlsx_path)
    worksheet = workbook.add_worksheet()

    # write column names on line 1
    workbook_index = 1

    # defining column names at index 1
    worksheet.write('A'+str(workbook_index), "COLUMN 1")
    worksheet.write('B'+str(workbook_index), "COLUMN 2")
    worksheet.write('C'+str(workbook_index), "COLUMN 3")
    
    workbook_index += 1

    # resize image to paste in xlsx
    image_width = 140.0
    image_height = 182.0
    cell_width = 20.0
    cell_height = 15.0

    x_scale = cell_width/image_width
    y_scale = cell_height/image_height

    for index, image_path in enumerate(glob(input_path+"/*")):
        filename = image_path.split("\\")[-1][:-4]

        # column one
        worksheet.write('A'+str(workbook_index), "text 1 column 1")

        # column two
        worksheet.write('B'+str(workbook_index), "text 1 column 2")

        # column three
        worksheet.insert_image('C'+str(workbook_index), image_path, {'x_scale': x_scale, 'y_scale': y_scale})
        
        # set next line or column number
        workbook_index+=6

    workbook.close()

# path to images folder
input_ = "G:\\codes\\scripts\\images"
# output will be saved in output folder
main(input_)

