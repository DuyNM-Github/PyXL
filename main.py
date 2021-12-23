import os
import datetime
from time import strftime

from library.ExcelHandler import ExcelHandler

if __name__ == '__main__':

    SAMPLE_FOLDER = "./sample/"

    excel = ExcelHandler()
    excel.open_excel(os.path.join(SAMPLE_FOLDER, "Financial Sample.xlsx"), alias="finance")
    excel.open_excel(os.path.join(SAMPLE_FOLDER, "sampledatafoodsales.xlsx"), alias="food")
    excel.open_excel(os.path.join(SAMPLE_FOLDER, "exceltables.xlsx"), alias="tables")
    print(excel.get_loaded_workbooks())

    excel.switch_workbook("food")
    excel.set_active_sheet("FoodSales")
    print(excel.get_sheet_value("E2"))
    print(excel.get_sheet_value("A3:H3"))
    print(excel.get_sheet_value("4:5"))
    print(excel.get_sheet_value("6"))

    excel.switch_workbook("tables")
    excel.set_active_sheet("OrdersTable")
    print(excel.get_all_tables())
    new_row = excel.extend_table_body("Orders", 1)
    print(new_row)
    excel.insert_data_to_row(new_row, [(datetime.datetime(2022, 5, 1, 0, 0).strftime("%m-%d-%Y")),
                                       'East', 'Paper', 73, 12.95, '?fn?', '?fn?', '?fn?', '=1'])
    print(excel.get_table_data("Orders"))

    excel.switch_workbook("food")
    excel.set_active_sheet("FoodSales")
    new_rows = excel.extend_table_body("Sales_Data", 2)
    print(new_rows)
    for new_row in new_rows:
        excel.insert_data_to_row(new_row, [(datetime.datetime(2022, 5, 1, 0, 0).strftime("%m-%d-%Y")),
                                           'East', 'Los Angeles', 'Cookies', 'Oatmeal', 50, 3.5, '?fn?'])
    print("Here: " + str(excel.get_table_data("Sales_Data")))

    excel.switch_workbook("finance")
    excel.copy_data("A2:B2", "test", full_col=True)
    excel.paste_data("Q2:R2", "test", full_col=True, overwrite=True)

    # excel.save_active_workbook()


