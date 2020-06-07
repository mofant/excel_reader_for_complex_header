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

    def get_same_value_type_count(self, col_types: list, col: "sheet_col"):
        """扫描一列的数据了下，获取每一行开始，其相同的数据类型数量
        回溯统计前面行中出现相同类型的次数。
        相同类型的定义是： 
            数值和字符串数值是相同的。
            空白符匹配所有的类型。
            字符串和空字符串是相同的。
        现在存在的问题是：
            1、空白的Unicode string应该与任何类型匹配。
            2、空字符串应该与任何类型匹配。
            3、空白的单元格应该任何类型匹配。

        Arguments:
            col_types {list} -- 一列的数据类型
            col {int}   --- sheet 的列号
        """
        row_map_same_value_count = {}
        for col_index, current_type in enumerate(col_types):
            row_map_same_value_count[col_index] = [current_type, 0]
            index = 0
            current_cell_value = self.sheet.cell(col_index, col).value
            while index < col_index:
                is_match = False
                # 空白的unicode 与任何类型匹配。
                temp_col_type_count = row_map_same_value_count[index]
                if current_type == 1 and not current_cell_value:
                    is_match = True
                # 空白和空字符串与任何类型匹配。
                elif current_type in [0, 6]:
                    is_match = True
                # 数字字符串和数字匹配
                elif current_type == 1 and temp_col_type_count[0] == 2:
                    if current_cell_value.isdigit():
                        is_match = True
                # 相等的时候
                elif temp_col_type_count[0] == current_type:
                    is_match = True
                # 如果当前值是字符串或者数值或者日期或者是boolean，匹配以前的空字符.
                elif current_type in [1, 2, 3, 4]:
                    if temp_col_type_count[0] in [0, 6]:
                        is_match = True
                    elif temp_col_type_count[0] == 1:
                        temp_col_value = self.sheet.cell(index, col).value
                        if not temp_col_value:
                            is_match = True

                if is_match:
                    temp_col_type_count[1] += 1
                index += 1

        return row_map_same_value_count

    def get_same_value_type_row(self):
        """获取行号，从那一行开始，所有列的数据都一致。

        Arguments:
            sheet {xlrd.sheet} -- excel sheet
        """
        # nrows = self.sheet.nrows
        ncols = self.sheet.ncols

        # 列方向的类型数据。
        sheet_col_types = [
            list(self.sheet.col_types(col)) for col in range(ncols)
        ]

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

        return col_same_type_count_list

    def _get_max_sub_continue_list(self, col_list: list) -> list:
        """寻找数组中连续最长的子串(倒序)

        Args:
            col_list (list): 列相同类型统计数组

        Returns:
            list: 连续最长子串的index list
        """
        col_continue_rows = []
        index = 0
        while index < len(col_list):
            temp = []
            temp.append(index)
            for j in range(index+1, len(col_list)):
                if j == len(col_list):
                    break
                pre = col_list[j]
                bef = col_list[j - 1]
                if bef - pre == 1:
                    temp.append(j)
                else:
                    break
            index = j
            col_continue_rows.append(temp)
            if j == len(col_list) - 1:
                break
        len_count = [len(each) for each in col_continue_rows]
        max_count = max(len_count)
        return col_continue_rows[len_count.index(max_count)]

    def search_max_same_type_rows(self) -> list:
        """查找最多连续相同数据类型的的行号数组。
        查找实现逻辑：
            1、获取每一列的的每个单元格的数据类型。
            2、在

        Returns:
            [int] -- 最多连续相同数据类型的的行信息
        """
        col_matrix = self.general_same_value_type_matrix()
        # 每一列中数据类型连续最多的开始行号和连续次数
        col_max_sub_con_rows = []
        for col_type_count in col_matrix:
            max_sub_con_rows = self._get_max_sub_continue_list(col_type_count)
            col_max_sub_con_rows.append(max_sub_con_rows)

        # 从每列连续的行号中，找到较多相同行号的。
        sub_con_lens = [len(each) for each in col_max_sub_con_rows]
        same_len_counter = Counter(sub_con_lens)
        most_same_len = same_len_counter.most_common()

        most_same_type_rows = []
        for con_len, _ in most_same_len:
            # 校验长度相同列的和是否都是一致的。如果一致，他就是最合适的连续相同列。
            col_max_sub_con_row_indexs = [index for index in range(
                len(sub_con_lens)) if sub_con_lens[index] == con_len]
            col_index_sum = [sum(col_max_sub_con_rows[index])
                             for index in col_max_sub_con_row_indexs]
            if len(set(col_index_sum)) == 1:
                most_same_type_rows = col_max_sub_con_rows[col_max_sub_con_row_indexs[0]]
                break
        return most_same_type_rows


class ColHeader:

    def __init__(self, sheet: xlrd.sheet, workbook: "xlrd.workbook", col_header_rows: list):
        self.sheet = sheet
        self.workbook = workbook
        self.nrow = sheet.nrows
        self.ncol = sheet.ncols
        self.col_header_rows = col_header_rows
        self.header_row = []  # 可能是大标题的行
        self.strip = True

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
            # 如果跟之前的col header重名，自动加上序号
            if col_header in col_headers:
                col_header += str(col)
            if self.strip:
                col_header = col_header.strip()
                col_header = col_header.replace('\n', "")
                col_header = col_header.replace(" ", "")
            col_headers.append(col_header)
        return col_headers


class ExcelCompxReader:
    """读取带有复杂表头的Excel的数据。
    """

    def __init__(self, file_path: str, strip: bool = True):
        self.file_path = file_path
        self.strip = strip

    def _open_workbook(self):
        try:
            self.workbook = xlrd.open_workbook(
                self.file_path, formatting_info=True)
        except Exception:
            raise

    def _get_continue_sub_list(self, data_list: list) -> list:
        """获取数组中值连续的子数组。

        Arguments:
            data_list {list} -- 已经排好序的数组

        Returns:
            list -- 连续值的子数组。
        """
        sub_lists = []
        index = 0
        while index < len(data_list):
            start_index = index
            for sec_index in range(start_index+1, len(data_list)):
                pre = data_list[sec_index]
                bef = data_list[sec_index - 1]
                # print(f"{pre} - {bef} = {pre - bef}")
                if pre - bef == 1:
                    continue
                else:
                    break
            if sec_index == len(data_list) - 1:
                sub_lists.append(data_list[start_index:])
            else:
                sub_lists.append(data_list[start_index: sec_index])
            if index == sec_index:
                index += 1
            else:
                index = sec_index
            if index == len(data_list) - 1:
                break
        return sub_lists

    def _get_continue_data_rows(self, data_rows: list) -> list:
        """获得连续的最长的数据行号

        Arguments:
            data_rows {list} -- 数据行

        Returns:
            list -- 数据行
        """
        sub_lists = self._get_continue_sub_list(data_rows)
        sub_lens = [len(sub) for sub in sub_lists]
        max_value = max(sub_lens)
        data_rows = sub_lists[sub_lens.index(max_value)]
        return data_rows

    def _split_col_header_and_data_row(self,
                                       no_merge_rows: list,
                                       containe_border_rows: list,
                                       same_type_rows: list) -> tuple:
        """
        表头与数据的划分。
        1、没有合并的单元格的行与每个单元格均由边框以及相同连续值的行的交集。
        2、在三种情况中出现两次及以上的行号。
        3、数据列一定是连续的行号，去除非连续的行号。
        """
        res_rows = no_merge_rows + containe_border_rows + same_type_rows
        row_count = Counter(res_rows)
        more_than_twich_row = [row for row,
                               count in row_count.items() if count >= 2]
        data_rows = sorted(more_than_twich_row)
        data_rows = self._get_continue_data_rows(data_rows)
        header_rows = [row for row in range(data_rows[0])]
        return data_rows, header_rows

    def _get_col_data(self, col_header_record: list, data_rows: list) -> dict:
        """获取excel中的每一列数据

        Returns:
            [dict] -- {"col_name": col_value_list}
        """
        res = {}
        for index, col_header in enumerate(col_header_record):

            col_value = [self.sheet.cell(
                row, index).value for row in data_rows]

            # 去除空的列
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


if __name__ == "__main__":
    file_path = "../test/excels/A1-01.xls"
    # file_path = "../test/excels/test_border.xls"
    # data = read_excel(file_path, "Sheet2")
    data = read_excel(file_path)
    print(data)
