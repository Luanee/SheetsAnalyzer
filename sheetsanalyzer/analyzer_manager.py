# !/usr/bin/python3.7
from os.path import isfile, join, isdir
from os import listdir, walk
from pathlib import Path
import pandas as pd
from xlsxwriter.utility import xl_cell_to_rowcol, xl_range, xl_rowcol_to_cell_fast


class FileManager:
    def __init__(self):
        super().__init__()
        self.basic_settings()
        self.setDelimiter(";")

    def basic_settings(self):
        self._path = None
        self._filetyp = None
        self._keywords = None
        self._origin_files = None
        self._range_text = None
        self._user_cell_range_text = None
        self._subfolder = False
        self._hidden_sheets = False
        self._all_sheets = True
        self._file_length = []
        self._cell_range = []
        self._user_cell_range = []
        self._sheet_names = []
        self._user_sheet_names = None
        self._file_delimiter = None

    def setPath(self, path):
        path = Path(path)
        if isdir(path):
            self._path = path
            self.setFiles()

    def getPath(self):
        return self._path

    def setFileTyp(self, type):
        self._filetyp = type
        self.setFiles()

    def getFileTyp(self):
        return self._filetyp

    def setSubFolder(self, mode):
        self._subfolder = mode
        self.setFiles()

    def getSubFolder(self):
        return self._subfolder

    def setKeywords(self, keys):
        if not self._origin_files:
            self.set_origin_files()

        self._keywords = keys
        self.set_key_files()

    def getKeywords(self):
        return self._keywords

    def setHiddenSheetsSettings(self, mode):
        self._hidden_sheets = mode

    def getHiddenSheetsSettings(self):
        return self._hidden_sheets

    def set_origin_files(self):
        self._origin_files = self._files

    def setFiles(self):
        self._files = []

        if self._path and self._filetyp and self._subfolder:
            for (dirpath, dirnames, filenames) in walk(self._path):
                self._files += [join(dirpath, file)
                                for file in filenames if file[-(len(file) - file.rfind(".") - 1):] == self._filetyp]

        elif self._path and self._subfolder:
            for (dirpath, dirnames, filenames) in walk(self._path):
                self._files += [join(dirpath, file) for file in filenames]

        elif self._path and self._filetyp:
            self._files = [join(self._path, f) for f in listdir(self._path) if isfile(
                join(self._path, f)) and f[-(len(f) - f.rfind(".") - 1):] == self._filetyp]

        elif self._path:
            self._files = [join(self._path, f) for f in listdir(self._path) if isfile(
                join(self._path, f))]

        if self._keywords:
            self.set_origin_files()
            self.set_key_files()

        self.setRange()

    def set_key_files(self):
        files = []

        for keyword in self._keywords:
            files += [f for f in self._origin_files if f.find(keyword) > -1]

        self._files = files

    def setDelimiter(self, delimiter):
        self._file_delimiter = delimiter

    def getDelimiter(self):
        return self._file_delimiter

    def setFileLength(self, column_number):
        self._file_length = column_number

    def getFileLength(self):
        return self._file_length

    def getFiles(self):
        return self._files

    def count_files(self):
        return len(self._files)

    def files_exist(self):
        if self.count_files() > 0:
            return True

        return False

    def get_file_length(self, file):
        largest_column_count = 0

        with open(file, 'r', encoding="ISO-8859-1") as temp_f:
            lines = temp_f.readlines()

            for l in lines:
                column_count = len(l.split(self._file_delimiter)) + 1

                largest_column_count = column_count if largest_column_count < column_count else largest_column_count

        # Close file
        temp_f.close()

        self.setFileLength([i for i in range(0, largest_column_count)])

    def setRange(self):
        if self.files_exist() and self._filetyp:
            self.get_file_length(self._files[0])

            if self._filetyp in ["xls", "xlsx", "xlsm"]:
                xl = pd.ExcelFile(self._files[0])
                df = xl.parse()
                self.setSheetNames(xl.sheet_names)
            else:
                df = pd.read_csv(self._files[0], names=self._file_length,
                                 engine="python", delimiter=self._file_delimiter,
                                 skip_blank_lines=False, skipinitialspace=True)

                if df.iloc[:, -1].isnull().values.all():
                    df.drop(df.columns[-1], inplace=True, axis=1)

            self.setRangeText(xl_range(0, 0, df.shape[0] - 1, df.shape[1] - 1))
            self._cell_range = (0, 0, df.shape[0] - 1, df.shape[1] - 1)
        else:
            self.setRangeText(" - not defined - ")

    def getRange(self):
        return self._cell_range

    def setRangeText(self, range_text):
        self._range_text = range_text

    def getRangeText(self):
        return self._range_text

    def setUserRange(self, range):
        self._user_cell_range = self.get_all_user_cells(range)
        self.setUserRangeText(self._user_cell_range[0] + ":" + self._user_cell_range[-1])

    def getUserRange(self):
        return self._user_cell_range

    def getUserCellRange(self):
        return [xl_cell_to_rowcol(cell) for cell in self._user_cell_range]

    def setUserRangeText(self, range_text):
        self._user_cell_range_text = range_text

    def getUserRangeText(self):
        return self._user_cell_range_text

    def check_user_cell_range(self):
        print(self._user_cell_range)
        return all([self.is_user_cell_in_range(xl_cell_to_rowcol(cell)) for cell in self._user_cell_range])

    def setAllSheets(self, mode):
        self._all_sheets = mode

    def getAllSheets(self):
        return self._all_sheets

    def setSheetNames(self, sheet_names):
        self._sheet_names = sheet_names

    def getSheetNames(self):
        return self._sheet_names

    def setUserSheetNames(self, sheet_names):
        self._user_sheet_names = []

        def getIndex(element):
            if checkNumber(element):
                return int(element) - 1
            elif checkString(element):
                return self._sheet_names.index(element)

            return None

        def checkNumber(number):
            if number.isdigit() and int(number) <= len(self._sheet_names):
                return True
            return False

        def checkString(name):
            if name in self._sheet_names:
                return True
            return False

        for element in sheet_names:
            if not element.find(":") > 0:
                index = getIndex(element)

                if index is not None:
                    self._user_sheet_names.append([index, True])
                else:
                    self._user_sheet_names.append([element, False])
            else:
                element = element.split(":")
                index_1 = getIndex(element[0])
                index_2 = getIndex(element[1])

                if index_1 is not None and index_2 is not None:
                    self._user_sheet_names.extend([[i, True] for i in range(index_1, index_2 + 1)])
                else:
                    self._user_sheet_names.append([":".join(element), False])

    def getUserSheets(self):
        return self._user_sheet_names

    def getUserSheetIndexes(self):
        if self._user_sheet_names:
            return [elem[0] for elem in self._user_sheet_names]

    def check_user_sheet_names(self):
        if self._user_sheet_names:
            return all([elem[1] for elem in self._user_sheet_names])
        return False

    def check_sheet_range(self):
        if self._all_sheets and self._hidden_sheets:
            return "All sheets (incl. hidden sheets)"
        elif self._all_sheets and not self._hidden_sheets:
            return "All sheets (excl. hidden sheets)"
        elif self.check_user_sheet_names():
            return [elem[0] for elem in self._user_sheet_names]

    def get_all_user_cells(self, cell_range_list):
        """
        Searches the cells to be merged for cell ranges and extracts the individual cells from these.

        keywords:
            cell_range_list -- list of cells and ranges (list)

        return:
            cell_list -- updated list of cells (list)
        """
        range_list = [elem for elem in cell_range_list if elem.find(":") > 0]
        cell_list = [elem for elem in cell_range_list if elem not in range_list]

        for j in range(len(range_list)):
            min_row, min_col, max_row, max_col = self.xl_range_reverse(range_list[j])

            for col in range(min_col, max_col + 1):
                for row in range(min_row, max_row + 1):
                    new_cell = xl_rowcol_to_cell_fast(row, col)
                    cell_list.append(new_cell)

        cell_list.sort()
        return cell_list

    def is_user_cell_in_range(self, cell):
        if self._cell_range[0] <= cell[0] <= self._cell_range[2] and self._cell_range[1] <= cell[1] <= self._cell_range[3]:
            return True

        return False

    def xl_range_reverse(self, range):
        cell = range.split(":")
        range = (xl_cell_to_rowcol(cell[0])[0], xl_cell_to_rowcol(cell[0])[1],
                 xl_cell_to_rowcol(cell[1])[0], xl_cell_to_rowcol(cell[1])[1])

        return self.control_range(range)

    def control_range(self, cell_range):
        min_row = min(cell_range[0], cell_range[2])
        min_col = min(cell_range[1], cell_range[3])
        max_row = max(cell_range[0], cell_range[2])
        max_col = max(cell_range[1], cell_range[3])

        return (min_row, min_col, max_row, max_col)

    def attributes(self):
        """
        Returns a string describing the filemanager attributes.

        :return: str
        """

        return ('(Filemanager) Attributes... \n'
                f'Folder Path:           {self._path}\n'
                f'File typ:              {self._filetyp}\n'
                f'Subfolder:             {self._subfolder}\n'
                f'Keywords:              {self._keywords}\n'
                f'Files Exist:           {self.files_exist()}\n'
                f'Number of Files:       {self.count_files()}\n'
                f'File Delimiter:        {self._file_delimiter}\n'
                f'File Range:            {self._range_text}\n'
                f'Number of RangeCells:  {len(self._cell_range)}\n'
                f'User Range:            {self._user_cell_range_text}\n'
                f'Number of RangeCells:  {len(self._user_cell_range)}\n'
                f'UserRange in Range:    {self.check_user_cell_range()}\n'
                f'Range of Sheets:       {self.check_sheet_range()}')

    def ready_to_run(self):
        if self._files and self.check_user_cell_range() and self._filetyp and self._all_sheets or self.check_user_sheet_names():
            return True

        return False
