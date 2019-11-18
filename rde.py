import os
import xlrd
from xlrd.sheet import Sheet
import xlsxwriter

HOME_PATH = os.path.dirname(os.path.abspath(__file__))


class RemoveDublicatesExcel(object):
    """
    An object that allows you to remove duplicates from the Excel by comparing
     a certain column according to the specified parameters
    """
    def __init__(self, source: str, src_home: bool = False,
                 out_home: bool = False, output: str = None,
                 name: bool = True, count: bool = True):
        self._source = f'{HOME_PATH}\\{source}' if src_home else source
        self._output = f'{HOME_PATH}\\{output}' if out_home else output
        self._sheet = 0
        self._comparison_col = 0
        self._start_row = 0
        self._dict_col = {}
        self._dict_col_category = {}
        self._join = False
        self._delimiter = '>'
        self._info_name = name
        self._count_name = count
        self._number_count = 0
        self.out_data = {}

    def sheet_index(self, index: int) -> None:
        """
        The transfer index of the book. Starts with 0

        :param index: Integer
        :return:
        """
        self._sheet = index

    def comparison_col_index(self, index: int) -> None:
        """
        Passing the index of the column on which to calculate the same records.
        :param index: Integer
        :return:
        """
        self._comparison_col = index

    def start_row(self, index: int) -> None:
        """
        Index of the line from which you want to start viewing the workbook
        :param index: Integer
        :return:
        """
        self._start_row = index

    def dict_view_col(self, dict_data: dict) -> None:
        """
        A dictionary with the name and indexes of the columns you want to
        view. Name - the key, the index - value.
        :param dict_data: Dictionary
        :return:
        """
        if isinstance(dict_data, dict):
            self._dict_col = dict_data
        else:
            print('Data not dictionary')
            raise ValueError(dict_data)

    def dict_category(self, dict_data: dict, join: bool = False) -> None:
        """
        Dictionary with title and indexes of columns which are categories.
        Name - the key, the index - value. Combine fields into one or not,
        specifies the boolean value
        :param dict_data: Dictionary
        :param join: True/False
        :return:
        """
        if isinstance(dict_data, dict):
            self._dict_col_category = dict_data
            self._join = join
        else:
            print('Data not dictionary')
            raise ValueError(dict_data)

    def set_delimiter_category(self, delimiter: str) -> None:
        """
        A string value that will divide the category name
        :param delimiter: String
        :return:
        """
        self._delimiter = delimiter

    def __bypass_for_spec_col(self, sheet: Sheet, row_num: int) -> None:
        """
        A private function that is executed when you specify a dictionary
        of certain fields.
        :param sheet: Book
        :param row_num: Number row
        :return:
        """
        _ = sheet.row_values(
            row_num, self._comparison_col, self._comparison_col+1
        )[0]

        if self.out_data.get(_):
            if self._info_name:
                print(_)
            self._number_count += 1
            return

        self.out_data[_] = {}

        for k, v in self._dict_col.items():
            value = sheet.row_values(row_num, v, v + 1)[0]
            self.out_data[_][k] = value

        if self._dict_col_category:
            self.__get_category(_, sheet, row_num)

    def __bypass_default(self, sheet: Sheet, row_num: int) -> None:
        """
        Private function, executed without specifying a dictionary of
         certain fields.
        :param sheet: Book
        :param row_num: Number row
        :return:
        """
        _ = self.out_data.get(sheet.row_values(
            row_num, self._comparison_col, self._comparison_col + 1
        ))[0]

        if _:
            if self._info_name:
                print(_)
            self._number_count += 1
            return

        self.out_data[_] = {}

        values = sheet.row_values(row_num)

        num = 0
        for val in values:
            self.out_data[_][num] = val
            num += 1

    def __get_category(self, key: str, sheet: Sheet, row_num: int) -> None:
        """
        The private function is executed if the dictionary of certain category
         fields is specified.
        :param key: Key record
        :param sheet: Book
        :param row_num: Number row
        :return:
        """
        if self._join:
            cat = ''
            for v in self._dict_col_category.values():
                _ = sheet.row_values(row_num, v, v + 1)[0]
                if cat:
                    cat += f'{self._delimiter}{_}'
                else:
                    cat = _
            self.out_data[key]['category'] = cat

        else:
            for k, v in self._dict_col_category:
                value = sheet.row_values(row_num, v, v+1)[0]
                self.out_data[key][k] = value

    def write_output(self, data=None) -> None:
        """
        A function that records the transferred data from the dictionary.
        The first line is written based on the dictionary keys.
        :param data: Dictionary data
        :return: Record XLS
        """
        data = self.out_data if not data else data
        wb = xlsxwriter.Workbook(self._output)
        ws = wb.add_worksheet()

        row = 0
        col = 0

        name = next(v for i, v in enumerate(data.values()) if i == 0)
        for key in name.keys():
            ws.write(row, col, key)
            col += 1

        col = 0
        row += 1

        for _ in data.values():
            for v in _.values():
                ws.write(row, col, v)
                col += 1
            col = 0
            row += 1

        wb.close()

    def exec(self) -> None:
        """
        The function of the next run. First, the file is read, then a new
        one is written.
        :return:
        """
        rb = xlrd.open_workbook(self._source, formatting_info=True)
        sheet = rb.sheet_by_index(self._sheet)

        for row_num in range(self._start_row, sheet.nrows):
            if self._dict_col:
                self.__bypass_for_spec_col(sheet, row_num)
            else:
                self.__bypass_default(sheet, row_num)

        if self._count_name:
            print(self._number_count)

        self.write_output(self.out_data)


# Example
rde = RemoveDublicatesExcel(
    source='src_file.xls', src_home=True,  out_home=True, output='out.xls'
)
rde.start_row(1)
rde.comparison_col_index(1)
rde.dict_view_col(
    {
        'name': 1,
        'prev_pic': 6,
        'detail_pic': 9,
        'prev_text': 7,
        'detail_text': 10
     }
)
rde.dict_category(
    {
        'cat_1': 27,
        'cat_2': 28,
        'cat_3': 29
    }, join=True
)
rde.exec()
