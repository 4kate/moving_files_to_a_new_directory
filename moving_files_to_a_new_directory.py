#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      kkurosad2
#
# Created:     24.10.2018
# Copyright:   (c) kkurosad2 2018
# Licence:     <your licence>
#-------------------------------------------------------------------------------
import os
import xlrd
import shutil


main_path = os.path.dirname(__file__)
new_directory = os.path.join(main_path,"new")
excel = os.path.join(main_path,"list_of_copies.xlsx")
excel_sheet = 'Sheet1'
errors_in_excel = os.path.join(main_path,"errors_in_excel.csv")


if not os.path.isdir(new_directory):
    os.mkdir(new_directory)

workbook = xlrd.open_workbook(excel)
sheet = workbook.sheet_by_name(excel_sheet)

with open(errors_in_excel,'wt') as file:
    for value in sheet.col_values(0):
        path_to_file = os.path.join(main_path,value)
        if os.path.isfile(path_to_file):
            shutil.move(path_to_file, new_directory)
        elif not os.path.isfile(path_to_file):
            file.write(value+"\n")
workbook.release_resources()
del workbook