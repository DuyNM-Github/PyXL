import os
import datetime
from time import strftime

from library.ExcelHandler import ExcelHandler

if __name__ == '__main__':

    SAMPLE_FOLDER = "./sample/"

    excel = ExcelHandler()
    excel.open_workbook(os.path.join(SAMPLE_FOLDER, "Financial Sample.xlsx"), alias="finance")
    excel.open_workbook(os.path.join(SAMPLE_FOLDER, "sampledatafoodsales.xlsx"), alias="food")
    excel.open_workbook(os.path.join(SAMPLE_FOLDER, "exceltables.xlsx"), alias="tables")
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
    new_rows = excel.extend_table_body("Sales_Data", 3)
    print(new_rows)
    for new_row in new_rows:
        excel.insert_data_to_row(new_row, [(datetime.datetime(2022, 5, 1, 0, 0).strftime("%m-%d-%Y")),
                                           'East', 'Los Angeles', 'Cookies', 'Oatmeal', 50, 3.5, '?fn?'])
    print("Here: " + str(excel.get_table_data("Sales_Data")))

    excel.switch_workbook("finance")
    excel.copy_data("A2:B5", "test", entire_col=False)
    excel.paste_data("Q2:R5", "test", entire_col=False, overwrite=True)

    excel.close_workbook(alias="finance")

    print(excel.get_active_workbook())

    excel.open_workbook(path_to_file="./sample/Financial Sample.xlsx")
    print(excel.get_active_workbook())
    print(excel.loaded_workbooks)
    excel.close_workbook(file_name="./sample/Financial Sample.xlsx")
    print(excel.get_active_workbook())
    excel.open_workbook(path_to_file="./sample/Financial Sample.xlsx", alias="finance")

    excel.open_workbook(path_to_file="sample/chart.xlsx", alias="chart")
    chart = excel.create_chart_col("A10", chart_title="Sample chart",
                                   chart_x_title="Sample Length (mm)", chart_y_title="Test number",
                                   reference_data_range="B1:C7", reference_category_range="A2:A7",
                                   chart_style=4)

    excel.copy_chart(chart, "testcopy")
    excel.paste_chart("testcopy", "K10")

    excel.save_active_workbook()
