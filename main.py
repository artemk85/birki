# -*- coding: utf-8 -*-
#
# Generate tags on product from excel
#
# Edit by Kochetkov Artem
# skype: artemk_85
# mail: kochetkov1985@mail.ru
#

from __future__ import absolute_import, division, print_function, unicode_literals
from io import BytesIO
import os

import upcean
import openpyxl
import xlsxwriter
from barcode import EAN13
from barcode.writer import ImageWriter


start_row_read = 16
end_row_read = 59
# col A - BE или 1 - 58


def get_data_from_xls(fn: str):
    data = []
    xls_file = os.path.abspath(fn)
    wb = openpyxl.load_workbook(filename=xls_file, read_only=True)
    ws = wb['Лист1']

    for row in range(start_row_read, end_row_read):
        data_row = []
        for col in range(1, 58):
            data_row.append(ws.cell(row=row, column=col).value)
        data.append(data_row)
    return data

all_file = 'all.xlsx'
barcode_file = 'barcode_all.xlsx'

xsl_data = get_data_from_xls(all_file)
# xsl_data = [[43, 180, 'Кольцо обручальное, БК, 18 (ш5)', 'СПЕЦ', 'DЕ', 'ЗОЛОТО', 585, '18,0', '125000', 4300128141, 'КРАСНЫЙ', 'БЕЗ ПОКРЫТИЯ', '2200001231238', 20320, 8000, 20320, 15749, None, '18,0', 585, None, 4300128141, '2200001231238', '05.06.2021', None, None, None, None, '13.04.2021', '4300013266', None, 48, 2.54, 2.54, 121.92, 121.92, 0, 'RUB', 0, 54.17, 'Г', 1, 0, 4318.29, 62.7, 0, 0, 6604.41, 1.5, 2.54, 2.54, 121.92, 121.92, 72.398, None, 'ОБРУЧ', 'БЕЗ АЛМ.ОГРАНКИ']]
# print(xsl_data)

doc = xlsxwriter.Workbook(filename=barcode_file)

for elem in xsl_data:
    doc_ws = doc.add_worksheet(str(elem[12]))
    doc_ws.set_column('A:T', 0.415)
    doc_ws.set_default_row(5.25)

    doc_ws.set_row_pixels(10, 11)
    doc_ws.set_row_pixels(11, 11)
    doc_ws.set_row_pixels(12, 9)
    doc_ws.set_row_pixels(13, 11)
    doc_ws.set_row_pixels(20, 10)
    doc_ws.set_row_pixels(21, 13)
    doc_ws.set_row_pixels(22, 9)
    doc_ws.set_row_pixels(23, 14)
    doc_ws.set_row_pixels(24, 14)
    doc_ws.set_row_pixels(25, 6)
    doc_ws.set_row_pixels(26, 4)
    doc_ws.set_row_pixels(27, 11)

    # doc_ws.print_area('A1:J20')

    # Наименование изделия
    name_format = doc.add_format(
        {
            'text_wrap': True,
            'font_name': 'Times New Roman',
            'font_size': 6,
            'text_h_align': 2,
            'text_v_align': 2,
        }
    )
    doc_ws.merge_range('B11:S12', elem[2], name_format)

    # Название товарного направления
    ntn_format = doc.add_format({
            'font_name': 'Arial',
            'font_size': 7,
            'text_h_align': 1,
            'text_v_align': 2,
        }
    )
    doc_ws.merge_range('B13:L13', elem[3], ntn_format)

    # Коллекция
    koll_format = doc.add_format({
            'font_name': 'Arial',
            'font_size': 7,
            'text_h_align': 3,
            'text_v_align': 2,
        }
    )
    doc_ws.merge_range('P13:S13', elem[4], koll_format)

    # Тип металла
    type_met_format = doc.add_format({
            'font_name': 'Times New Roman',
            'font_size': 5,
            'text_h_align': 2,
            'text_v_align': 2,
        }
    )
    doc_ws.merge_range('B14:H14', elem[5], type_met_format)

    # Проба
    proba_format = doc.add_format({
            'font_name': 'Times New Roman',
            'font_size': 8,
            'text_h_align': 2,
            'text_v_align': 2,
        }
    )
    doc_ws.merge_range('I14:L14', elem[6], proba_format)

    # Размер
    raz_format = doc.add_format({
            'font_name': 'Times New Roman',
            'font_size': 5,
            'text_h_align': 2,
            'text_v_align': 2,
        }
    )
    doc_ws.merge_range('M14:O14', 'Р-р', raz_format)

    raz1_format = doc.add_format({
            'font_name': 'Times New Roman',
            'font_size': 8,
            'text_h_align': 2,
            'text_v_align': 2,
        }
    )
    doc_ws.merge_range('P14:S14', elem[7], raz1_format)

    # Артикул
    art_format = doc.add_format({
            'font_name': 'Times New Roman',
            'font_size': 5,
            'text_h_align': 1,
            'text_v_align': 2,
        }
    )
    doc_ws.merge_range('B15:S15', f'Арт {elem[8]}', art_format)

    # САП код
    sap_format = doc.add_format({
            'font_name': 'Times New Roman',
            'font_size': 5,
            'text_h_align': 1,  # left
            'text_v_align': 2,
        }
    )
    doc_ws.merge_range('B16:S16', f'САП код: {elem[9]}', sap_format)

    # Цвет металла
    cvet_format = doc.add_format({
            'font_name': 'Times New Roman',
            'font_size': 5,
            'text_h_align': 3,  # right
            'text_v_align': 2,
        }
    )
    doc_ws.merge_range('B17:S17', f'Цв. {elem[10]}', cvet_format)

    # Покрытие


    # Штрихкод
    image_data = BytesIO()
    ean = EAN13(str(elem[12]), writer=ImageWriter())
    ean.write(image_data, options={"write_text": False})

    x_scale = 0.191  # *13.89
    y_scale = 0.115  # *5.39

    doc_ws.insert_image(
        'A18',
        str(elem[12]),
        {
            'image_data': image_data,
            'x_offset': 0,
            'y_offset': 0,
            'x_scale': x_scale,
            'y_scale': y_scale,
            'object_position': 0,
        }
    )

    bc = doc.add_format({
            'font_name': 'Arial',
            'font_size': 7,
            'text_h_align': 2,
            'text_v_align': 1,  # top
        }
    )
    doc_ws.merge_range('B21:S21', str(elem[12]), bc)

    # Цена
    cena_format = doc.add_format({
            'font_name': 'Times New Roman',
            'font_size': 8,
            'text_h_align': 2,
            'text_v_align': 2,
        }
    )
    doc_ws.merge_range('B22:F22', 'Цена:', cena_format)

    cena2_format = doc.add_format({
            'bold': True,
            'font_name': 'Times New Roman',
            'font_size': 10,
            'text_h_align': 2,
            'text_v_align': 2,
        }
    )
    doc_ws.merge_range('G22:O22', elem[13], cena2_format)

    cena3_format = doc.add_format({
            'font_name': 'Times New Roman',
            'font_size': 8,
            'text_h_align': 2,
            'text_v_align': 2,
        }
    )
    doc_ws.merge_range('P22:S22', 'руб.', cena3_format)

    # Цена за грамм
    cena_format = doc.add_format({
            'font_name': 'Times New Roman',
            'font_size': 5,
            'text_h_align': 2,
            'text_v_align': 1,
        }
    )
    doc_ws.merge_range('B23:F23', 'За гр.:', cena_format)

    cena2_format = doc.add_format({
            'font_name': 'Times New Roman',
            'font_size': 5,
            'text_h_align': 2,
            'text_v_align': 1,
        }
    )
    doc_ws.merge_range('G23:O23', elem[14], cena2_format)

    cena3_format = doc.add_format({
            'font_name': 'Times New Roman',
            'font_size': 5,
            'text_h_align': 2,
            'text_v_align': 1,
        }
    )
    doc_ws.merge_range('P23:S23', 'руб.', cena3_format)

    # Размер
    raz1_format = doc.add_format({
            'font_name': 'Arial',
            'font_size': 7,
            'text_h_align': 2,
            'text_v_align': 2,
        }
    )
    doc_ws.merge_range('B24:F24', elem[18], raz1_format)

    # Цена прочерк
    cena_format = doc.add_format({
            'font_name': 'Times New Roman',
            'font_size': 9,
            'text_h_align': 2,
            'text_v_align': 2,
            'font_strikeout': 1,
        }
    )
    doc_ws.merge_range('L24:S24', f'{elem[15]}p', cena_format)

    # Проба
    if elem[19] == 375:
        proba_format = doc.add_format({
                'font_name': 'Arial',
                'font_size': 7,
                'text_h_align': 2,
                'text_v_align': 2,
            }
        )
        doc_ws.merge_range('B25:F25', elem[19], proba_format)

    # Цена продаж
    cena_format = doc.add_format({
            'font_name': 'Times New Roman',
            'font_size': 9,
            'text_h_align': 2,
            'text_v_align': 2,
        }
    )
    doc_ws.merge_range('L25:S25', f'{elem[16]}p', cena_format)

    # Штрихкод
    image_data2 = BytesIO()
    ean2 = EAN13(str(elem[22]), writer=ImageWriter())
    ean2.write(image_data2, options={"write_text": False, "includetext": True})

    x_scale = 0.191  # *13.89
    y_scale = 0.057  # *5.39

    doc_ws.insert_image(
        'A26',
        f'{elem[22]}.png',
        {
            'image_data': image_data2,
            'x_offset': 0,
            'y_offset': 0,
            'x_scale': x_scale,
            'y_scale': y_scale,
            'object_position': 0,
        }
    )

    # САП код
    sap_format = doc.add_format({
            'font_name': 'Arial',
            'font_size': 8,
            'text_h_align': 2,
            'text_v_align': 2,
        }
    )
    doc_ws.merge_range('B28:S28', elem[9], sap_format)

doc.close()
