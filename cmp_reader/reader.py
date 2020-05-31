import xlrd
from xlrd.sheet import Cell
from xlrd import sheet
from collections import Counter
from functools import reduce
import math

__all__ = ["ExcelCompxReader", "read_excel"]

MAX_SEARCH_ROW = 100


class SheetIndex:
    """sheet Index查找
    假定每个excel都是有index的。
    index的在表格中的定义是：
        1、在列中合并单元格，其合并的单元格值都一样。
        2、index列在excel列的最开始的地方。
        3、
    Returns:
        [type] -- [description]
    """

    def __init__(self, sheet: xlrd.sheet, workbook: "xlrd.workbook"):
        self.sheet = sheet
        self.workbook = workbook
        self.nrows = sheet.nrows
        self.ncols = sheet.ncols


class SheetMerged:
    """sheet中的合并单元格
    """

    def __init__(self, sheet):
        self.sheet = sheet

    def get_merge_rows(self) -> list:
        """获取有合并的行号列表

        Arguments:
            sheet {xrld.sheet} -- 

        Returns:
            list -- 有合并单元格行号的list
        """
        merge_cells = self.sheet.merged_cells
        merge_rows = [each for merge_cell in merge_cells for each in range(
            merge_cell[0], merge_cell[1])]
        merge_rows = list(set(merge_rows))
        return merge_rows

    def get_no_merge_rows(self) -> list:
        """获取没有合并单元格的行号z

        Returns:
            list -- 没有合并单元格的行号
        """
        rows = [row for row in range(self.sheet.nrows)]
        merge_rows = self.get_merge_rows()
        no_merge_row = set(rows) - set(merge_rows)
        return list(no_merge_row)


class SheetBorder:
    """sheet边框相关类
    """

    def __init__(self, sheet: xlrd.sheet, workbook: "xlrd.Workbook"):
        self.sheet = sheet
        self.workbook = workbook
        self.nrows = sheet.nrows
        self.ncols = sheet.ncols

    def get_cell_XF(self, cell) -> "XF":
        """获取单元格的XF

        Arguments:
            cell {sheet.cell} -- 单元格
            workbook {xlrd.open_workbook} -- 打开的workbook

        Returns:
            [XF] -- [formatting info]]
        """
        xf_index = cell.xf_index
        return self.workbook.xf_list[xf_index]

    def get_sheet_border(self) -> "Matrix":
        """获取所有单元格都有内边框的最开始的行号.
            当数据量比较大的时候，检测到有连续5行
            由于存在单元格只有3边有单元格。

        Arguments:
            sheet {xlrd.sheet} -- [description]
        """

        sheet_border_info = []
        for row in range(self.nrows):
            row_border = []
            for col in range(self.ncols):
                xf = self.get_cell_XF(self.sheet.cell(row, col))
                line_style = []
                line_style.append(1 if xf.border.top_line_style else 0)
                line_style.append(1 if xf.border.bottom_line_style else 0)
                line_style.append(1 if xf.border.left_line_style else 0)
                line_style.append(1 if xf.border.right_line_style else 0)
                row_border.append(sum(line_style))
            sheet_border_info.append(row_border)
        return sheet_border_info

    def get_each_cell_has_border_rows(self) -> list:
        """获取每个单元格均有边框的行号。
        每个单元格均由单元格的判定标准是：超过3边有边框。
        在一行中，如果超过75%的单元格有边框，那就是有边框的行。

        Returns:
            [list] -- list(行号)
        """
        sheet_border_info = self.get_sheet_border()

        border_rows = []
        per_75_row_num = math.ceil(self.ncols * 0.8)
        for row in range(self.nrows):
            has_border_cells = [
                each for each in sheet_border_info[row] if each >= 3]
            if len(has_border_cells) >= per_75_row_num:
                border_rows.append(row)
        return border_rows


class SheetType:
    """搜索每一列中数据类型相同的单元格。
    缺： 单元格为空的时候。
    """

    def __init__(self, sheet: xlrd.sheet):
        self.sheet = sheet
        self.nrow = sheet.nrows
        self.ncol = sheet.ncols

    def chain_type_count(self, row_map_same_value_count: dict, current_type: int, last_index: int):
        """链式反应，查找上一行的数据类型是否跟当前的相同，如果相同，计数加一，并递归往上查找。
            递归慢，丢弃
        Arguments:
            row_map_same_value_count {dict} -- 每行中连续相同值的计数
            current_type {int} -- 查找行当前的类型
            last_index {int} -- 上一行的index
        """
        if last_index < 0:
            return
        last = row_map_same_value_count[last_index]
        # 非空的不相同类型也返回
        if last[0] != current_type and current_type != 6:
            return
        last[1] += 1
        row_map_same_value_count[last_index] = last
        self.chain_type_count(row_map_same_value_count,
                              current_type, last_index - 1)

    def process_pre_same_type(self,
                              row_map_same_value_count: dict,
                              current_type: int,
                              current_row: int,
                              col: int):
        """回溯统计前面行中出现相同类型的次数。
        相同类型的定义是： 
            数值和字符串数值是相同的。
            空白符匹配所有的类型。
            字符串和空字符串是相同的。


        Arguments:
            row_map_same_value_cout {dict} -- 行号与自此行开始相同的次数
            current_type {int} -- 当前统计的数据类型。
            current_row {int} -- 当前的行
            col: 当前值的列
        """
        last_row = current_row - 1

        while last_row >= 0:
            is_match = False
            last = row_map_same_value_count[last_row]

            if last[0] == 2 and current_type == 1:
                cell_value = self.sheet.cell(current_row, col).value
                if cell_value.isdigit():
                    is_match = True
            elif current_type == 6:
                is_match = True
            elif last[0] == 1 and current_type == 0:
                is_match = True
            elif last[0] == current_type:
                is_match = True

            if not is_match:
                return

            last[1] += 1
            row_map_same_value_count[last_row] = last
            # row_map_same_value_count[last_row] = row_map_same_value_count[last_row] + 1
            last_row -= 1

    def get_same_value_type_count(self, col_types: list, col: "sheet_col"):
        """扫描一列的数据了下，获取每一行开始，其相同的数据类型数量

        Arguments:
            col_types {list} -- 一列的数据类型
        """
        row_map_same_value_count = {}
        for col_index, ctype in enumerate(col_types):
            row_map_same_value_count[col_index] = [ctype, 0]
            self.process_pre_same_type(
                row_map_same_value_count, ctype, col_index, col)
            # self.chain_type_count(row_map_same_value_count, ctype, col_index-1)
        return row_map_same_value_count

    def get_same_value_type_row(self):
        """获取行号，从那一行开始，所有列的数据都一致。

        Arguments:
            sheet {xlrd.sheet} -- excel sheet
        """
        nrows = self.sheet.nrows
        ncols = self.sheet.ncols

        sheet_col_types = [
            [self.sheet.cell(i, j).ctype for i in range(nrows)] for j in range(ncols)]
        row_map_same_value_count_list = []
        for col in range(ncols):
            row_map_same_value_count_list.append(
                self.get_same_value_type_count(sheet_col_types[col], col))
        return row_map_same_value_count_list

    def general_same_value_type_matrix(self):
        """构建连续相同类型的矩阵

        Arguments:
        """
        row_map_same_value_count_list = self.get_same_value_type_row()
        col_same_type_count_list = [
            [value[1] for _, value in col.items()]
            for col in row_map_same_value_count_list
        ]
        matrix = []
        for row in range(self.nrow):
            row_value = []
            for col in range(self.ncol):
                row_value.append(col_same_type_count_list[col][row])
            matrix.append(row_value)
        return matrix

    def search_max_same_type_rows(self) -> list:
        """查找最多连续相同数据类型的的行号。

        Returns:
            [int] -- 最多连续相同数据类型的开始行号
        """
        matrix = self.general_same_value_type_matrix()
        # 每一列中数据类型连续最多的开始行号和连续次数
        col_max_continue_type_rows = []
        for col in range(self.ncol):
            col_value = [[matrix[row][col], row] for row in range(self.nrow)]
            max_count_row = max(col_value)
            col_max_continue_type_rows.append(
                {
                    row for row in range(max_count_row[1], sum(max_count_row) + 1)
                }
            )
        return list(reduce(lambda x, y: x & y, col_max_continue_type_rows))


class ColHeader:

    def __init__(self, sheet: xlrd.sheet, workbook: "xlrd.workbook", col_header_rows: list):
        self.sheet = sheet
        self.workbook = workbook
        self.nrow = sheet.nrows
        self.ncol = sheet.ncols
        self.col_header_rows = col_header_rows
        self.header_row = []  # 可能是大标题的行

    def get_cell_XF(self, cell) -> "XF":
        """获取单元格的XF

        Arguments:
            cell {sheet.cell} -- 单元格
            workbook {xlrd.open_workbook} -- 打开的workbook

        Returns:
            [XF] -- [formatting info]]
        """
        xf_index = cell.xf_index
        return self.workbook.xf_list[xf_index]

    def get_merged_info(self) -> list:
        """获取合并的单元格
        排除掉疑似表头的合并项。
        疑似表头的合并项的条件是：
            合并单元格长度等于数据列长度。

        Returns:
            list -- [description]
        """
        merge_cells = self.sheet.merged_cells
        no_header_merge_cells = []
        for r_start, r_end, c_start, c_end in merge_cells:
            if c_end - c_start == self.ncol:
                self.header_row.append(r_start)
            else:
                no_header_merge_cells.append(
                    (r_start, r_end, c_start, c_end)
                )
        return no_header_merge_cells

    def get_merge_col_header(self, merge_cell: tuple, T_col_header: dict) -> dict:
        """获取一个合并单元格中，其顶行的所有单元格的值.
        并且根据t型合并的数据，赋予其横向的字段值。

        Arguments:
            merge_cell {tuple} -- 一个合并单元格的信息
            T_col_header {dict} -- t型横向字段值

        Returns:
            dict -- {"row_col": value}
        """
        value = self.sheet.cell(merge_cell[0], merge_cell[2]).value
        row_col_values = {}
        start_row = merge_cell[0]
        t_col_key = f"{start_row}_{merge_cell[2]}"
        if t_col_key in T_col_header:
            value = T_col_header[t_col_key] + value
        for col in range(merge_cell[2], merge_cell[3]):
            row_col_values[f"{start_row}_{col}"] = str(value)
        return row_col_values

    def parse_T_col_header(self, merge_cells: list):
        """解析T型的合并单元格表头。

        Arguments:
            merge_cells {list} -- 合并的单元格信息
        """
        t_col_header_value = {}
        for r_start, _, c_start, c_end in merge_cells:
            if r_start in self.col_header_rows and c_end < self.ncol:
                left_cell_xf = self.get_cell_XF(
                    self.sheet.cell(r_start, c_start))
                right_cell_xf = self.get_cell_XF(
                    self.sheet.cell(r_start, c_end))
                if left_cell_xf.border.right_line_style == right_cell_xf.border.left_line_style == 0:
                    t_col_key = f"{r_start}_{c_end}"
                    t_col_header_value[t_col_key] = self.sheet.cell(
                        r_start, c_start).value
        return t_col_header_value

    def search_col_header(self):
        merge_cell_col_values = {}
        merge_cells = self.get_merged_info()

        t_col_headers = self.parse_T_col_header(merge_cells)

        for merge_cell in merge_cells:
            merge_cell_col_values.update(
                self.get_merge_col_header(merge_cell, t_col_headers))

        col_header_row = list(set(self.col_header_rows) - set(self.header_row))
        col_headers = []

        for col in range(self.ncol):
            col_header = ""
            # 每一列的值等于其表头行中所有值的拼接。
            for row in col_header_row:
                merge_key = f"{row}_{col}"
                t_col_value = ""
                merge_col_value = ""
                cell_value = self.sheet.cell(row, col).value
                if merge_key in t_col_headers:
                    t_col_value = t_col_headers[merge_key]

                if merge_key in merge_cell_col_values and cell_value != merge_cell_col_values[merge_key]:
                    merge_col_value = merge_cell_col_values[merge_key]
                # 一个字段的值是由： 单元格自身的值、合并单元格中空缺的值、t型对齐的值三个决定。
                # 如果合并单元格中存在值，那么这行的值就是合并单元格的值，且合并单元格的值和单元格的值以及t的值一致的时候。
                # 如果合并单元格存在值，其值也不等于单元格自身的值和t型对齐，则当前行的值为单元格的值加上合并单元格的值。
                if merge_col_value:
                    cell_value = str(merge_col_value)
                # 如果不存在合并单元格，其值将有t型和自身值决定。
                else:
                    cell_value = str(t_col_value) + str(cell_value)
                
                col_header += cell_value
            #如果跟之前的col header重名，自动加上序号
            if col_header in col_headers:
                col_header += str(col)
            col_headers.append(col_header)
        return col_headers


class ExcelCompxReader:
    """读取带有复杂表头的Excel的数据。
    """
    def __init__(self, file_path: str):
        self.file_path = file_path

    def _open_workbook(self):
        try:
            self.workbook = xlrd.open_workbook(self.file_path, formatting_info=True)
        except Exception:
            raise

    def _split_col_header_and_data_row(self,
                                       no_merge_rows: list,
                                       containe_border_rows: list,
                                       same_type_rows: list) -> tuple:
        """
        表头与数据的划分。
        1、没有合并的单元格的行与每个单元格均由边框以及相同连续值的行的交集。
        2、在三种情况中出现两次及以上的行号。
        """
        res_rows = no_merge_rows + containe_border_rows + same_type_rows
        row_count = Counter(res_rows)
        more_than_twich_row = [row for row,
                               count in row_count.items() if count >= 2]
        data_rows = sorted(more_than_twich_row)
        header_rows = [row for row in range(data_rows[0])]
        return data_rows, header_rows

    def _get_col_data(self, col_header_record: list, data_rows: list) -> dict:
        """获取excel中的每一列数据

        Returns:
            [dict] -- {"col_name": col_value_list}
        """
        res = {}
        for index, col_header in enumerate(col_header_record):
            # 去除空的列
            col_value = [self.sheet.cell(
                row, index).value for row in data_rows]
            if not col_header:
                if not all(col_value):
                    continue
            res[col_header] = col_value
        return res

    def read_excel(self, sheet_name=None) -> dict:
        """读取excel中的数据

        Keyword Arguments:
            sheet_name {str} -- sheet name (default: {None})

        Returns:
            dict -- {"col_name": col_value_list}
        """        
        self._open_workbook()
        if sheet_name:
            self.sheet = self.workbook.sheet_by_name(sheet_name)
        else:
            self.sheet = self.workbook.sheet_by_index(0)

        sheet_merger = SheetMerged(self.sheet)
        sheet_border = SheetBorder(self.sheet, self.workbook)
        sheet_typer = SheetType(self.sheet)

        no_merge_rows = sheet_merger.get_no_merge_rows()
        containe_border_rows = sheet_border.get_each_cell_has_border_rows()
        same_type_rows = sheet_typer.search_max_same_type_rows()

        data_rows, header_rows = self._split_col_header_and_data_row(
            no_merge_rows, containe_border_rows, same_type_rows)

        col_header = ColHeader(self.sheet, self.workbook, header_rows)
        col_header_record = col_header.search_col_header()

        return self._get_col_data(col_header_record, data_rows)

def read_excel(filename, sheet_name=None) -> dict:
    """从文件中读取一张带有复杂表头的sheet

    Arguments:
        filename {str} -- excel file path 

    Keyword Arguments:
        sheet_name {str} -- sheet name (default: {None})

    Returns:
        dict -- {"col_header": value_list}
    """    
    reader = ExcelCompxReader(filename)
    return reader.read_excel(sheet_name)
