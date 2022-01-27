import qrcode
import PIL
from docx import Document
from docx.shared import Mm
import os
from docx.shared import Cm, Inches
import excel


PIC_FILE_NAME = "pic.png"
NUMBER_COLS = 3
NUMBER_ROWS = 0
TABLE_CELL_WIDTH = Cm(5)
WIDTH_QR = Cm(3)
HEIGHT_QR = Cm(3)

path_excel_file = "Test table.xlsx"
list_num_col_for_QR =   (1, 2, 3, 4, 6, 7, 8, 9)
list_num_col_QR_descr = (4, 7, 8, 9)
path_doc = "QR_codes.docx"


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


# функция создает таблицу в формате ".doc" с QR-кодами
# и описанием мат. технических ресурсов
def make_QR_doc_list(object_excel, object_doc, list_num_col_for_QR: list,
                     list_num_col_QR_descr: list, path_doc: str):

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


    table = doc_QR_document.add_table(NUMBER_ROWS, NUMBER_COLS)
    table.style ='Table Grid'
    column_num_doc_table = 0
    row_cells = table.add_row().cells

    for row_num_doc_table in range(1, 25): #object_excel.rows + 1): FIXME
        print(row_num_doc_table) # !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        make_img_QR(object_excel, row_num_doc_table + 1, list_num_col_for_QR)
        paragraph = row_cells[column_num_doc_table].paragraphs[0]
        run = paragraph.add_run()
        row_cells[column_num_doc_table].add_paragraph().add_run().\
            add_text(get_string_from_excel(object_excel, row_num_doc_table + 1, list_num_col_QR_descr))
        if row_num_doc_table % 3 == 0 and row_num_doc_table != object_excel.rows - 1:
            row_cells = table.add_row().cells       # new string
        run.add_picture(PIC_FILE_NAME, width = WIDTH_QR, height = HEIGHT_QR)
        column_num_doc_table = row_num_doc_table % 3


    os.remove(PIC_FILE_NAME)

    doc_QR_document.save("test_doc.docx")


def get_string_from_excel(excel_file: excel.Excel, num_row: int, list_num_col: list):
    data = ""
    for i in list_num_col:
        data += excel_file.get_data_cell(num_row, i)
    return data


def make_img_QR(excel_file, num_row, list_num_col_for_QR):
    str_QR = get_string_from_excel(excel_file, num_row, list_num_col_for_QR)
    img = qrcode.make(str_QR)
    img.save(PIC_FILE_NAME)


# main
ex = excel.Excel(path_excel_file)
make_QR_doc_list(ex, "!!!!!!!!!!!!!!!!!!!!", list_num_col_for_QR, list_num_col_QR_descr, path_doc)


"""
    table.autofit = False
    for i in range(3):
        for cell in table.columns[i].cells:
            cell.width = TABLE_CELL_WIDTH
            cell.left_margin = 0
            cell.top_margin = 0
            cell.right_margin = 0


    sections = doc_QR_document.sections
    margin = Cm(0.5)

    for section in sections:
        section.top_margin = margin
        #section.bottom_margin = margin
        section.left_margin = margin
        #section.right_margin = margin
        """
