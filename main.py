# -*- coding: utf-8 -*-
#
# Generate tags on product from excel
#
# Edit by Kochetkov Artem
# skype: artemk_85
# mail: kochetkov1985@mail.ru
#

from io import BytesIO
import os

import barcode
import openpyxl
import xlsxwriter
from barcode import EAN13
from barcode.writer import ImageWriter

start_row_read = 16
end_row_read = 59
# col A - BE или 1 - 58

def get_data_from_xls():
    data = []
    xls_file = os.path.abspath('all.xlsx')
    wb = openpyxl.load_workbook(filename=xls_file, read_only=True)
    ws = wb['Лист1']

    for row in range(start_row_read, end_row_read):
        data_row = []
        for col in range(1, 58):
            data_row.append(ws.cell(row=row, column=col).value)
        data.append(data_row)
    return data


barcode_file = 'barcode_all.xlsx'

#xsl_data = get_data_from_xls()
xsl_data = [[1, 10, 'Серьги-пуссеты, БК', 'БК', 'Ч4', 'ЗОЛОТО', 375, '19.0', 'С20ГМ0012А0-ЗЛ31-70', 4300127675, 'КРАСНЫЙ', 'БЕЗ ПОКРЫТИЯ', '2200001226579', 5990, 5500.46, 5990, 4580, None, None, 375, None, 4300127675, '2200001226579', '08.06.2021', None, None, None, None, '13.04.2021', '4300013233', None, 70, 1.089, 1.089, 76.23, 76.23, 0, 'RUB', 0, 158.33, 'Г', 1, 0, 4318.29, 62.7, 0, 0, 12069.5, 3, 1.089, 1.089, 76.23, 76.23, 45.928, None, 'ОВАЛ', 'БЕЗ АЛМ.ОГРАНКИ']]
#print(xsl_data)

doc = xlsxwriter.Workbook(filename=barcode_file)

for elem in xsl_data:
    doc_ws = doc.add_worksheet(str(elem[12]))
    doc_ws.set_page_view()
    doc_ws.set_column('A:J', 0.715)
    doc_ws.set_default_row(7.20)
    doc_ws.print_area('A1:J20')

    # Наименование изделия
    name_format = doc.add_format()
    name_format.set_font_name('Times New Roman')
    name_format.set_font_size(6)
    name_format.set_align('center')
    name_format.set_align('vcenter')
    doc_ws.merge_range('A5:J6', elem[2], name_format)

    # Название товарного направления
    ntn_format = doc.add_format()
    ntn_format.set_font_name('Arial')
    ntn_format.set_font_size(7)
    ntn_format.set_align('center')
    ntn_format.set_align('vcenter')
    doc_ws.merge_range('A7:C7', elem[3], ntn_format)

    # Коллекция
    koll_format = doc.add_format()
    koll_format.set_font_name('Arial')
    koll_format.set_font_size(7)
    koll_format.set_align('center')
    koll_format.set_align('vcenter')
    doc_ws.merge_range('H7:J7', elem[4], koll_format)

    # Тип металла
    type_met_format = doc.add_format()
    type_met_format.set_font_name('Times New Roman')
    type_met_format.set_font_size(5)
    type_met_format.set_align('center')
    type_met_format.set_align('vcenter')
    doc_ws.merge_range('A8:C8', elem[5], type_met_format)

    # Проба
    proba_format = doc.add_format()
    proba_format.set_font_name('Times New Roman')
    proba_format.set_font_size(6)
    proba_format.set_align('center')
    proba_format.set_align('vcenter')
    doc_ws.merge_range('D8:F8', elem[6], proba_format)

    # Размер
    raz_format = doc.add_format()
    raz_format.set_font_name('Times New Roman')
    raz_format.set_font_size(5)
    raz_format.set_align('center')
    raz_format.set_align('vcenter')
    doc_ws.write('G8', 'Р-р', raz_format)

    raz1_format = doc.add_format()
    raz1_format.set_font_name('Times New Roman')
    raz1_format.set_font_size(8)
    raz1_format.set_align('center')
    raz1_format.set_align('vcenter')
    doc_ws.merge_range('H8:J8', elem[7], raz1_format)

    # Артикул
    art_format = doc.add_format()
    art_format.set_font_name('Times New Roman')
    art_format.set_font_size(5)
    art_format.set_align('left')
    art_format.set_align('vcenter')
    doc_ws.merge_range('A9:J9', elem[8], art_format)

    # САП код
    sap_format = doc.add_format()
    sap_format.set_font_name('Times New Roman')
    sap_format.set_font_size(5)
    sap_format.set_align('left')
    sap_format.set_align('vcenter')
    doc_ws.merge_range('A10:J10', elem[9], sap_format)

    # Цвет металла
    cvet_format = doc.add_format()
    cvet_format.set_font_name('Times New Roman')
    cvet_format.set_font_size(5)
    cvet_format.set_align('right')
    cvet_format.set_align('vcenter')
    doc_ws.merge_range('A11:J11', elem[10], cvet_format)

    # Покрытие


    # Штрихкод
    image_data = BytesIO()
    cod = barcode.get('upca', str(elem[12]))
    ean = EAN13(str(elem[12]), writer=ImageWriter())
    ean.write(image_data, options={"write_text": False})

    image_width = 143.78
    image_height = 79.01

    cell_width = 24.85
    cell_height = 7.67

    x_scale = cell_width / image_width
    y_scale = cell_height / image_height

    doc_ws.insert_image(
        'A12',
        f'{elem[12]}.png',
        {
            'image_data': image_data,
            'x_offset': 0,
            'y_offset': 0,
            'x_scale': x_scale,
            'y_scale': y_scale,
            'object_position': 0,
        }
    )

doc.close()
