import openpyxl
import tkinter as tk
from tkinter import filedialog


class ExcelTranslate(object):
    def __init__(self):
        pass

    @staticmethod
    def get_excel():
        main_excel = None
        root = tk.Tk()
        root.withdraw()
        file_path = filedialog.askopenfilename()
        if file_path:
            wb = openpyxl.load_workbook(file_path)
            main_excel = wb.active
        return main_excel

    @staticmethod
    def make_out_excel():
        # 新建一个表格
        out_excel = openpyxl.Workbook()
        # 选择活动工作表
        sheet = out_excel.active
        return out_excel, sheet

    def translate_en_to_chinese(self, word):
        pass

    def get_translate_write_excel_value(self):
        in_excel = self.get_excel()
        trans_excel_value = []
        for hang in range(2, 335):
            for lie in range(1, 3):
                cell = in_excel[hang][lie]
                # cell_value = cell.value
                # trans_value = self.translate_en_to_chinese(cell_value)
                trans_excel_value.append(cell.value)

        out_excel, sheet = self.make_out_excel()

        i = 0
        for hang_out in range(2, 335):
            for lie_out in 'BC':
                cell_name = str(lie_out) + str(hang_out)
                cell = sheet[cell_name]
                cell.value = trans_excel_value[i]
                i += 1
        out_excel.save('NASA_Image_Mars_Translate.xlsx')


if __name__ == '__main__':
    ExcelTranslate().get_translate_write_excel_value()
