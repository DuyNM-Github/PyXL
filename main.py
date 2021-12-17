import os

from library.ExcelHandler import ExcelHandler

if __name__ == '__main__':

    SAMPLE_FOLDER = "./sample/"

    excel = ExcelHandler()
    excel.open_excel(os.path.join(SAMPLE_FOLDER, "sampledatafoodsales.xlsx"))
    excel.open_excel(os.path.join(SAMPLE_FOLDER, "Financial Sample.xlsx"))
    excel.switch_excel_file("sampledatafoodsales.xlsx")
    print(excel.set_active_sheet("FoodSales"))

