# coding: utf-8

import logging
import yaml
import os
import string
from openpyxl import Workbook, load_workbook
import xlrd

logging.basicConfig(level=logging.DEBUG)

BLANK_LINES = 3


class ExcelDriver(object):
    def __init__(self):
        self.excel_type = 0

    def load_file(self, filename):
        pass

    def change_sheet(self):
        pass

    def get_cell_value(self):
        pass


class ExcelBook(object):
    def __init__(self):
        self.parse_mode = 1
        self.wb = None
        self.sheets = []
        self.ws = None
        self.cols = []
        self.row_range = []
        self.headers = []

    def load(self, xlsx_file_name="", yaml_config_file=""):
        if not xlsx_file_name:
            logging.warning("Missing xlsx file name.")
            return
        self.wb = load_workbook(xlsx_file_name)
        self.sheets = self.wb.sheetnames
        self.ws = self.wb.active
        # Load desc yaml config
        if not yaml_config_file:
            yaml_filename = xlsx_file_name.split('.')[0] + '.yml'
        else:
            yaml_filename = yaml_config_file
        skip_level, skip_table_headers = 0, 1
        if os.path.exists(yaml_filename):
            logging.info("find {}, read config".format(yaml_filename))
            data = {}
            with open(yaml_filename, encoding='UTF-8') as fp:
                data = yaml.load(fp, yaml.FullLoader)
                logging.debug("the YAML config is {}".format(data))
            headers = data.get('headers', {})
            skip_level = headers.get('skip_level', 0)
            skip_table_headers = headers.get('total_high', 1)
        # peek the rows and cols at the excel
        _, _, self.cols = self.get_cols_range(skip_level=skip_level)
        row_start, row_end, _ = self.get_rows_range(skip_table_headers=skip_table_headers)
        self.row_range = (row_start, row_end)
        # try to load headers
        self.headers = self.get_headers(headers_row=2)

    def get_cols_range(self, skip_level=0):
        """
            excel cols range peek
        :return:
        """
        logging.debug("get cols range")
        col_start = string.ascii_uppercase[0]
        col_end = ''
        col_array = []
        blank_col_count = 0
        for alphabet in string.ascii_uppercase:
            cell_pos = '{}{}'.format(alphabet, 1+skip_level)
            if self.ws[cell_pos].value:
                blank_col_count = 0
                col_array.append(alphabet)
                col_end = alphabet
                continue
            # blank col, start counting
            blank_col_count += 1
            if blank_col_count == 3:
                break
        logging.debug("col_start {}, col_end {}, col_array {}".format(col_start, col_end, col_array))
        return col_start, col_end, col_array

    def get_rows_range(self, skip_table_headers=1, the_first_col='A'):
        """
            excel rows range peek
            default table header rows set as one.
        :return:
        """
        logging.debug("getting rows range from excel, skip table headers {}, and the first col is {}".format(
            skip_table_headers, the_first_col
        ))
        row_start = 1 + skip_table_headers
        row_end = 1 + skip_table_headers
        row_array = []
        blank_row_count = 0
        while 1:
            pos_tag = '{}{}'.format(the_first_col, row_end)
            if self.ws[pos_tag].value:
                blank_row_count = 0
                row_array.append(pos_tag)
                row_end += 1
                continue
            # blank row, start counting
            blank_row_count += 1
            if blank_row_count == 3:
                row_end -= 1
                break
        logging.debug("row start {}, row end {}".format(row_start, row_end))
        return row_start, row_end, row_array

    def get_headers(self, headers_row=1, headers_blocks=[]):
        headers = []
        # _, _, col_range = self.get_cols_range()
        for col in self.cols:
            header_pos = "{}{}".format(col, headers_row)
            cell_value = self.ws[header_pos].value
            if not cell_value:
                continue
            headers.append(cell_value)
        logging.info("this excel file headers {}".format(headers))
        return headers

    def get_data(self):
        data = []
        row_start, row_end = self.row_range
        for i in range(row_start, row_end+1):
            single_data = {}
            for k in self.cols:
                pos = "{}{}".format(k, i)
                cell_value = self.ws[pos].value or ""
                t_h = self.headers[self.cols.index(k)]
                single_data[t_h] = cell_value
            data.append(single_data)
        return data


class SimpleExcelBook(ExcelBook):
    def __init__(self):
        super(SimpleExcelBook, self).__init__()


if __name__ == '__main__':
    book = SimpleExcelBook()
    # book.load('EnglishWordSmaple.xlsx')
    book.load('sample2.xls')
    data = book.get_data()
    print(data)

