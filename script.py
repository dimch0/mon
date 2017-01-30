#!/usr/bin/env python2
# -*- coding: utf-8 -*-


# https://github.com/mikemaccana/python-docx
# http://stackoverflow.com/questions/22169325/read-excel-file-in-python
# http://stackoverflow.com/questions/3198765/how-to-write-russian-characters-in-file


# The above encoding declaration is required and the file must be saved as UTF-8

from __future__ import with_statement   # Not required in Python 2.6 any more

import codecs
import xlsxwriter
from xlrd import open_workbook


ID_CARD = [u'паспорт', u'л.к.']
ISSUED_ON = [u'изд. На', u'издаден на', u'изд. на']
ISSUED_BY = u'МВР'

# passport = passport.encode('utf-8')
for i in ID_CARD:
    print i

for i in ISSUED_ON:
    print i
print "\n\n\n"



FOR_AGREEMENT_FILE = 'for-agreements.xlsx'

wb = open_workbook(FOR_AGREEMENT_FILE)
list_to_write = []


for sheet in wb.sheets():
    number_of_rows = sheet.nrows
    number_of_columns = sheet.ncols

    items = []

    rows = []
    for row in range(1, number_of_rows):
        for col in range(number_of_columns):

            value  = (sheet.cell(row,col).value)
            print "row:", row, "col:", col, "val:", value.split(',')
            v = value.split(',')
            list_to_write.append(v)
            # for item in v:
            #     print item
            #     list_to_write.append(item)
                # for p in ID_CARD:
                #     if p in item:
                #         print item
                # for i in ISSUED_ON:
                #     if i in item:
                #         print item



updated_file = "MM.xlsx"
workbook = xlsxwriter.Workbook(updated_file)
worksheet = workbook.add_worksheet()

for row in range(0, len(list_to_write)):
    for col in range(0, len(list_to_write[row])):
        worksheet.write(row, col, list_to_write[row][col])

workbook.close()

