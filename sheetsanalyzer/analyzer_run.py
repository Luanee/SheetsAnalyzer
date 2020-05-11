# !/usr/bin/python3.7
from PyQt5.QtCore import QThread, pyqtSignal
from pathlib import Path
import pandas as pd
from os.path import isfile, join, commonpath
from time import strftime

# TODO: create table format


class Analyzer(QThread):
    countChanged = pyqtSignal(int, int)

    def __init__(self, files, file_typ, file_delimiter, keywords, columns, user_range, cell_indexes, all_sheets, hidden_sheets, user_sheets, parent=None):
        super(Analyzer, self).__init__(parent)
        self._files = files
        self._save_filetyp = file_typ
        self._file_delimiter = file_delimiter
        self._keywords = keywords
        self._user_range = user_range
        self._cell_indexes = cell_indexes
        self._all_sheets = all_sheets
        self._hidden_sheets = hidden_sheets
        self._user_sheets = user_sheets
        self._file_length = columns
        self._sheets_name = "Sheet1"
        self._user_df = None
        self._writer = None
        self._percent = 0
        self._max_len = len(self._files)
        self._steps = 100 / self._max_len

    def get_dataframe(self, file):
        if file.rsplit(".")[-1] in ["xls", "xlsx", "xlsm"]:
            xl = pd.ExcelFile(file)
            return xl.parse()
        else:
            return pd.read_csv(file, names=self._file_length,
                               engine="python", delimiter=self._file_delimiter,
                               skip_blank_lines=False, skipinitialspace=True)

    def save_dataframe(self):
        self._save_path = Path(commonpath(self._files)).parent
        self._save_name = self.generate_savename()

        self._writer = pd.ExcelWriter(Path(self._save_path, self._save_name), engine='xlsxwriter')
        self._user_df.to_excel(self._writer, sheet_name=self._sheets_name, index=False)

        # self.format_result_file()

        self._writer.save()

    def generate_savename(self):
        save_name = strftime("%y%m%d") + "_SheetsAnalyzer." + self._save_filetyp

        index = 0
        while isfile(join(self._save_path, save_name)):
            index += 1
            save_name = strftime("%y%m%d") + "_SheetAnalyzer_" + \
                str(index) + "." + self._save_filetyp

        return save_name

    def format_result_file(self):
        workbook = self._writer.book
        worksheet = self._writer.sheets[self._sheets_name]

        format1 = workbook.add_format({'num_format': '#,##0.00'})
        worksheet.set_column('B:B', 18, format1)

    def run(self):
        columns = self._user_range.copy()
        columns.insert(0, "FileName")
        self._user_df = pd.DataFrame(columns=columns)

        for file in self._files:
            df = self.get_dataframe(file)

            extract_list = [df.iloc[cell[0], cell[1]] for cell in self._cell_indexes]
            self._user_df.loc[self._files.index(file)] = [Path(file).name] + extract_list

            self._percent += self._steps
            self.countChanged.emit(int(self._percent), self._files.index(file) + 1)

        self.save_dataframe()
