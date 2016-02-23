#!/usr/bin/python3.4
import xlwt
__author__ = 'Daniel Pasacrita'
__date__ = '2/23/16'


def output_to_excel_file(four_urls_unique, directory, filedate):
    """
    This will take the four_urls_unique param and output it to an excel spreadsheet.
    :param four_urls_unique: A set of every 404 returned from Elasticsearch
    :param directory: The present working directory
    :param filedate: The date the file was created.
    :return: N/A
    """
    # Output to an excel spreadsheet
    # Create the workbook
    book = xlwt.Workbook(encoding="utf-8")
    sheet1 = book.add_sheet("Sheet 1")
    # This will make the first column much larger, about 150 characters wide:
    first_col = sheet1.col(0)
    first_col.width = 256*150
    # Fill in the workbook
    i = 0
    for url in four_urls_unique:
        sheet1.write(i, 0, url)
        i += 1
    # Save the workbook
    book.save(directory+"/reports/"+"urls_spreadsheet."+filedate+".xls")

if __name__ == "__main__":
    print("Functions are fun!!")
