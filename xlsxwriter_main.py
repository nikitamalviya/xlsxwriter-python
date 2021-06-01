# from posix import XATTR_REPLACE
import xlsxwriter, os, sys
from glob import glob
from datetime import datetime

path = "G:\\codes\\scripts\\"
image_folder = path +"images"

def main(input_path, filename_split_keyword):
    # filename_split_keyword "\\"
    # print(" : ", input_path.split(filename_split_keyword)[-1])
    xlsx_path = "output\\"+  input_path.split(filename_split_keyword)[-1] + ".xlsx"

    # print("xlsx_path : ", xlsx_path)

    if os.path.exists(xlsx_path):
        os.remove(xlsx_path)
        print("Removing old XLSX file...")
    
    # creating file obj and file
    workbook = xlsxwriter.Workbook(xlsx_path)
    worksheet = workbook.add_worksheet()

    workbook_index = 1

    # defining column names at index 1
    worksheet.write('A'+str(workbook_index), "COLUMN 1")
    worksheet.write('B'+str(workbook_index), "COLUMN 2")
    worksheet.write('C'+str(workbook_index), "COLUMN 3")
    
    workbook_index += 1

    print(" : ", filename_split_keyword, xlsx_path)
    # exit()

    image_width = 140.0
    image_height = 182.0

    cell_width = 20.0
    cell_height = 15.0

    x_scale = cell_width/image_width
    y_scale = cell_height/image_height

    for index, image_path in enumerate(glob(input_path+"/*")):
        print("index : ", index)

        filename = image_path.split(filename_split_keyword)[-1][:-4] #.replace("\\","_")[:-4]
        print(image_path, image_path.split(filename_split_keyword), "filename : ", filename)

        # column one
        worksheet.write('A'+str(workbook_index), "text 1 column 1")

        # column two
        worksheet.write('B'+str(workbook_index), "text 1 column 2")

        # column three
        worksheet.insert_image('C'+str(workbook_index), image_path, {'x_scale': x_scale, 'y_scale': y_scale})
        
        # next line or more than that
        workbook_index+=6

    workbook.close()


input_ = "G:\\codes\\scripts\\images"
filename_split_keyword="scripts\\"
main(input_, filename_split_keyword)