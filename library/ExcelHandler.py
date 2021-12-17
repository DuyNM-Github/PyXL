import os

from openpyxl import Workbook, load_workbook
from openpyxl.utils.exceptions import InvalidFileException


class ExcelHandler:
    """
        Reusable chunk of codes that focus on the handling of Excel files
    """

    wb = Workbook()

    def __init__(self):
        self.active_excel_file = None
        self.loaded_workbooks = {}

    def open_excel(self, path_to_file):
        try:
            temp = load_workbook(path_to_file)
            self.active_excel_file = temp
            self.loaded_workbooks[os.path.basename(path_to_file)] = temp
            print(f"Excel file {os.path.basename(path_to_file)} loaded")
        except (InvalidFileException, FileNotFoundError) as error:
            print(f"FILE LOADING ERROR: {error}")

    def switch_excel_file(self, file_name):
        for key, values in self.loaded_workbooks.items():
            if key == file_name:
                self.active_excel_file = values
                print(f"Active workbook switched to {key}")
                return
        print("WORKBOOK SWITCHING ERROR: Requested workbook is not found")
        print(f"Hint. Loaded workbooks: {self.loaded_workbooks.keys()}")

    def set_active_sheet(self, sheet_name=None):
        wb = self.active_excel_file
        if wb is None:
            print("EXCEL ERROR: No Excel file is currently loaded")
            return None
        if sheet_name is not None:
            for s in range(len(wb.sheetnames)):
                if wb.sheetnames[s] == sheet_name:
                    wb.active = s
                    active_sheet = wb.active
                    print(f"Active sheet switched to {sheet_name}")
                    return active_sheet
            print(f"EXCEL ERROR: Sheet {sheet_name} not found")
            return None

    def read_sheet(self):
        wb = self.active_excel_file

