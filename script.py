#!/usr/bin/env python2
# -*- coding: utf-8 -*-

# http://stackoverflow.com/questions/10366596/how-to-read-contents-of-an-table-in-ms-word-file-using-python
# https://python-docx.readthedocs.io/en/latest/search.html?q=table&check_keywords=yes&area=default


from docx import Document
from docx.shared import Pt
from xlrd import open_workbook
from docx.enum.style import WD_STYLE_TYPE

FOR_AGREEMENT_FILE = 'for-agreements.xlsx'




DOG_TEMPLATE = "dog.docx"
DOG_NEW = u"Договор-за-обучение_BG-EN_Appendix-B_HRC-Foundation-002_{0}.docx"

name_bg = "NAME_BG_HERE"
name_e = "NAME_E_HERE"
info_1 = u"с ЕГН  NATIONAL_ID_HERE, л. к. ID_CARD_NO_HERE, изд. на ISSUED_ON_HERE от МВР – ISSUED_BY_HERE"
info_2 = 'ADDRESS_BG_HERE'
info_3 = "with National ID No. NATIONAL_ID_HERE, personal ID card No. ID_CARD_NO, issued on ISSUED_ON_HERE by the Ministry of Interior - ISSUED_BY_HERE,"
info_4 = "ADDRESS_E_HERE"


def write_doc(doc_template, new_doc, info_list):
    doc = Document(doc_template)
    table = doc.tables[0]

    style = doc.styles['Normal']
    font = style.font
    font.size = Pt(11)
    font.name = 'Cambria'

    style2 = doc.styles.add_style('Undefined', WD_STYLE_TYPE.PARAGRAPH)
    font2 = style2.font
    font2.size = Pt(11)
    font2.name = 'Cambria'

    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                if name_bg in paragraph.text:
                    paragraph.text = paragraph.text.replace(name_bg, "")
                    paragraph.add_run(info_list[0]).bold = True
                    paragraph.style = doc.styles['Normal']
                if name_e in paragraph.text:
                    paragraph.text = paragraph.text.replace(name_e, "")
                    paragraph.add_run(info_list[1]).bold = True
                    paragraph.style = doc.styles['Normal']
                if info_1 in paragraph.text:
                    paragraph.text = paragraph.text.replace(info_1, info_list[2])
                    paragraph.style = doc.styles['Undefined']
                if info_2 in paragraph.text:
                    paragraph.text = paragraph.text.replace(info_2, info_list[3])
                    paragraph.style = doc.styles['Undefined']
                if info_3 in paragraph.text:
                    paragraph.text = paragraph.text.replace(info_3, info_list[4])
                    paragraph.style = doc.styles['Undefined']
                if info_4 in paragraph.text:
                    paragraph.text = paragraph.text.replace(info_4, info_list[5])
                    paragraph.style = doc.styles['Undefined']
    doc.save(new_doc)




split_1 = u'Адрес по регистрация'
split_2 = u'registered address'


wb = open_workbook(FOR_AGREEMENT_FILE)


for sheet in wb.sheets():
    number_of_rows = sheet.nrows
    number_of_columns = sheet.ncols



    for row in range(1, number_of_rows):


        for col in range(number_of_columns):

            value  = (sheet.cell(row,col).value)

            if col == 0:
                v = value.split(split_1)
                name1 = v[0].split(',')[0]
                info1 = v[0].replace(name1+",", "")
                info2 = split_1 + v[1]
                # print "row:", row, "col:", col, "val:", v[0], v[1]

            if col == 1:
                v2 = value.split(split_2)
                name2 = v2[0].split(',')[0]

                info3 = v2[0].replace(name2+",", "")
                info4 = "Registered address" + v2[1]

        info_list = [name1, name2, info1, info2, info3, info4]
        new_doc_name = DOG_NEW.format("_".join(name2.split()))
        print new_doc_name
        write_doc(DOG_TEMPLATE, new_doc_name, info_list)


