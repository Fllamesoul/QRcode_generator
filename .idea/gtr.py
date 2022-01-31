import qrcode
import PIL
from docx import Document
from docx import table
from docx.shared import Mm
import os
from docx.shared import Cm, Inches
import imp
import openpyxl as xl


fp, pathname, description = imp.find_module("excel")
_mod = imp.load_module("excel", fp, pathname, description)

PIC_FILE_NAME = "pic.png"
NUMBER_COLS = 5 * 2
NUMBER_ROWS = 0
TABLE_CELL_WIDTH = Cm(5)
WIDTH_QR = Cm(1)
HEIGHT_QR = Cm(1)

path_excel_file = "Test table.xlsx"
list_num_col_for_QR =   (1, 2, 3, 4, 6, 7, 8, 9)
list_num_col_QR_descr = (4, 7, 8, 9)
path_doc = "QR_codes.docx"

"""
# создает объект ДОКУМЕНТ,
#   устанавливает поля = 0
#   устанавливает выравнивание по середине
#   создает таблицу с заданным количеством столбоцв
def create_doc_template(path_doc_file, num_rows, num_columns):
    doc_QR_document = Document()
    section = doc_QR_document.sections[0]
    section.page_height = Cm(29.7)      # высота листа в сантиметрах
    section.page_width = Cm(21.0)       # ширина листа в сантиметрах
    section.left_margin = Mm(0)         # левое поле в миллиметрах
    section.right_margin = Mm(0)        # правое поле в миллиметрах
    section.top_margin = Mm(0)         # верхнее поле в миллиметрах
    section.bottom_margin = Mm(0)      # нижнее поле в миллиметрах
    section.header_distance = Mm(0)    # отступ от верхнего края страницы до
                                        # нижнего края нижнего колонтитула
    section.footer_distance = Mm(0)    # отступ от нижнего края страницы до
                                        # нижнего края нижнего колонтитула

    table = doc_QR_document.add_table(num_rows, num_columns, )
    table.style ='Table Grid'
    table.autofit = False
    return doc_QR_document
    """


# функция создает таблицу в формате ".doc" с QR-кодами
# и описанием мат. технических ресурсов
def make_QR_doc_list(object_excel, rezult_doc_path, list_col_for_QR,
                     list_col_QR_descr, path_template):

    doc_QR_document = Document(path_template)
    table = doc_QR_document.tables[0]

    num_row_doc = 0
    num_col_doc = 0

    for count_excel in range(25): #FIXME 25!!!
        paragraph = table.row_cells[count_excel].paragraphs[0]
        if num_col_doc % 2:




    os.remove(PIC_FILE_NAME)

    doc_QR_document.save(rezult_doc_path)

    """
    
    for row_num_doc_table in range(25): #object_excel.rows + 1): FIXME
        print("num_row", num_row )
        print("num_col", num_col)
        print()

        make_img_QR(object_excel, row_num_doc_table + 1, list_num_col_for_QR)
        paragraph = row_cells[column_num_doc_table].paragraphs[0]
    #    paragraph.alignment = 1

        if row_num_doc_table % 2:
            run = paragraph.add_run()
            paragraph.add_picture(PIC_FILE_NAME, width = WIDTH_QR, height = HEIGHT_QR)
            num_row += 1
        #else:
            #cell = table.cell(num_row-1, num_col)
            #cell.text = "sdasd"
        num_col = row_num_doc_table % NUMBER_COLS

        print("num_row", num_row )
        print("num_col", num_col)
        print(row_num_doc_table) # !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        print()
        print()
        print()
        
            
     #   else:
      #      row_cells[column_num_doc_table].add_paragraph().add_run().\
      #          add_text(get_string_from_excel(object_excel, row_num_doc_table + 1, list_num_col_QR_descr))
   #     row_cells[column_num_doc_table].paragraphs[1].alignment = 1
   

        if row_num_doc_table % NUMBER_COLS == 0 and row_num_doc_table != object_excel.rows - 1:
            row_cells = table.add_row().cells       # new string

        
        column_num_doc_table = row_num_doc_table % NUMBER_COLS
"""


def get_string_from_excel(excel_file: _mod.Excel, num_row: int, list_num_col: list):
    data = ""
    for i in list_num_col:
        data += excel_file.get_data_cell(num_row, i)
    return data


def make_img_QR(excel_file, num_row, list_num_col_for_QR):
    str_QR = get_string_from_excel(excel_file, num_row, list_num_col_for_QR)
    img = qrcode.make(str_QR)
    img.save(PIC_FILE_NAME)


def _init_():
    fp, pathname, description = imp.find_module("excel")
    _mod = imp.load_module("excel", fp, pathname, description)
    #path_excel_file = input("Open excel file: (Для инвентаризации.xlsx by default)")
    path_excel_file = "Для инвентаризации.xlsx" #FIXME
    if (path_excel_file == ""):
        path_excel_file = "Для инвентаризации.xlsx"
    ex = _mod.Excel(path_excel_file)
    #path_template_docx_file = input("Open template docx file: (Template.docx by default)")
    path_template_docx_file = "Template.docx" #FIXME
    if (path_template_docx_file == ""):
        path_template_docx_file = "Template.docx"
    make_QR_doc_list(ex, "result.docx", list_num_col_for_QR, list_num_col_QR_descr, path_template_docx_file)


