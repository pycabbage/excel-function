#!/usr/bin/env python3
#coding: utf-8

from string import ascii_uppercase
import re
from typing import List


class Ref:
    def __init__(self, ref: str, type="auto", column_limit="", row_limit="") -> None:
        R = r"(\$?\d+)"
        C = r"(\$?[A-Za-z]+)"
        self.__EXCEL_DEFAULT_PATTERN = r"^((.+)!)?(" + \
            C+"|"+R+"|"+C+R+")(:("+C+"|"+R+"|"+C+R+"))?$"
        # self.__SPREADSHEET_DEFAULT_PATTERN = r"^('(.+)'!)?("+C+"|"+R+"|"+C+R+")(:("+C+"|"+R+"|"+C+R+"))?$"
        """
        2: sheetname
        4: column
        5: row
        6: column, 7: row
        10: column_end
        11: row_end
        12: column_end, 13: row_end
        """
        self.__PATTERNS = {
            "RANGE":             C+R+":"+C+R,
            "ALONE":             C+R,
            "ROW_RANGE":         R+":"+R,
            "COLUMN_RANGE":      C+":"+C,
            "ROW_RANGE_MORE":    C+R+":"+R,
            "COLUMN_RANGE_MORE": C+R+":"+C,
        }
        self.inputref = ref
        self.spectype = type  # auto, excel, spreadsheet
        self.type: str = "unparsed"  # range, alone, {range,alone}_more
        """
        range: ex. A1:B2
        alone: ex. A1
        {row,column}_range: ex. 1:2, A:B
        {row,column}_{range,alone}_more: ex. A1:A, A1:B, A1:2
        Note: "_more" is not supported on excel, so it will converted to :
          A1:A -> A1:A1048576
          A1:1 -> A1:XFD1
        """

        self.sheet: str = ""  # None to current, str to other

        self.column: str = ""  # A, B, C, ...
        self.row: str = ""  # 1, 2, 3, ...
        self.column_end: str = ""
        # A, B, C, ...
        # rc range:
        self.row_end: str = ""
        # 1, 2, 3, ...

        # set fixed
        self.column_fixed: bool = False
        self.row_fixed: bool = False
        self.column_end_fixed: bool = False
        self.row_end_fixed: bool = False

        self.width: int = 0
        self.height: int = 0

        # set spectype
        if self.spectype == "auto":
            if re.match(r"^'.+'!.+$", ref):
                self.spectype = "spreadsheet"
                self.column_limit = column_limit or "Z"
                self.row_limit = row_limit or "1000"
            else:
                self.spectype = "excel"
                self.column_limit = column_limit or "XFD"
                self.row_limit = row_limit or "1048576"

        if self.spectype == "excel":
            self.__parse_excel()
        elif self.spectype == "spreadsheet":
            self.__parse_spreadsheet()
        self.__lint()

    def __isfixed(self, rc: str) -> bool:
        return bool(re.search(r"^\$$", rc))

    def __lint(self):
        self.column = self.column.upper()
        self.column_end = self.column_end.upper()
        self.row = self.row.upper()
        self.row_end = self.row_end.upper()

    def __parse_excel(self) -> None:
        self.__parse_spreadsheet()
        if self.sheet:
            self.sheet = re.search("'(.+)'", self.sheet).group(1)
        if not self.column_end == self.row_end:
            if not self.column_end:
                self.column_end = self.column_limit
            elif not self.row_end:
                self.row_end = self.row_limit

        self.column_fixed = self.__isfixed(self.column)
        self.row_fixed = self.__isfixed(self.row)
        self.column_end_fixed = self.__isfixed(self.column_end)
        self.row_end_fixed = self.__isfixed(self.row_end)
        self.column = re.sub(r"^\$", "", self.column)
        self.row = re.sub(r"^\$", "", self.row)
        self.column_end = re.sub(r"^\$", "", self.column_end)
        self.row_end = re.sub(r"^\$", "", self.row_end)

    def __parse_spreadsheet(self) -> None:
        # get order by regex
        s = re.search(self.__EXCEL_DEFAULT_PATTERN, self.inputref)
        if not s:
            raise ValueError("Invalid reference: " + self.inputref)
        self.sheet = s.group(2)
        self.column = s.group(4) or s.group(6) or "A"
        self.row = s.group(5) or s.group(7) or "1"
        self.column_end = s.group(10) or s.group(12)  # or self.column
        self.row_end = s.group(11) or s.group(13)  # or self.row
        if self.column_end is self.row_end is None:
            self.column_end = self.column
            self.row_end = self.row
        elif self.column_end is not self.row_end:
            if self.column_end is None:
                self.column_end = ""
            elif self.row_end is None:
                self.row_end = ""

        # set fixed, removed fixed string
        self.column_fixed = self.__isfixed(self.column)
        self.row_fixed = self.__isfixed(self.row)
        self.column_end_fixed = self.__isfixed(self.column_end)
        self.row_end_fixed = self.__isfixed(self.row_end)
        self.column = re.sub(r"^\$", "", self.column)
        self.row = re.sub(r"^\$", "", self.row)
        self.column_end = re.sub(r"^\$", "", self.column_end)
        self.row_end = re.sub(r"^\$", "", self.row_end)

        # Fix reverse order
        if self.column and self.column_end and self.__column_to_number(self.column) > self.__column_to_number(
                self.column_end):
            self.column, self.column_end = self.column_end, self.column
        if type(self.row) is type(self.row_end) is int and int(self.row) > int(self.row_end):
            self.row, self.row_end = self.row_end, self.row

        if not self.column and not self.row:
            raise ValueError("Invalid reference: " + self.inputref)

    def __column_to_number(self, column: str, start: int = 1) -> int:
        number = start + ascii_uppercase.index(column[-1].upper())
        if len(column) > 1:
            number += self.__column_to_number(
                column[:-1], number) * len(ascii_uppercase)
        return number

    def move(self, pos: List[int]) -> None:
        """
        move to pos
        :param pos:
        :return:
        """
        pass

    def __str__(self):
        string = ""
        if self.sheet:
            if self.spectype == "excel":
                string += self.sheet + "!"
            else:
                string += "'" + self.sheet + "'!"
        string += self.column + self.row
        if (self.column_end or self.row_end) and self.column != self.column_end and self.row != self.row_end:
            string += ":"
            if self.column_end:
                string += self.column_end
            if self.row_end:
                string += self.row_end
        return string


if __name__ == "__main__":
    print(Ref("A1", type="spreadsheet"))
    print(Ref("B2", type="spreadsheet"))
    print(Ref("A1:B2", type="spreadsheet"))
    print(Ref("B:B", type="spreadsheet"))
    print(Ref("1:1", type="spreadsheet"))
