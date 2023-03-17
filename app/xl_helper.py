"""
openpyxlのヘルパー関数
"""


from copy import copy
from typing import Iterable
from openpyxl.cell.cell import Cell
from openpyxl.worksheet.header_footer import _HeaderFooterPart
from openpyxl.worksheet.worksheet import Worksheet


def copy_style(src: Cell, dst: Cell):
    """
    `src`のセルのスタイルで`dst`のスタイルを上書きする
    """
    dst.font = copy(src.font)
    dst.border = copy(src.border)
    dst.fill = copy(src.fill)
    dst.number_format = copy(src.number_format)
    dst.protection = copy(src.protection)
    dst.alignment = copy(src.alignment)


def all_header_footer_parts(ws: Worksheet) -> Iterable[_HeaderFooterPart]:
    wsprops = [a + b for
               a in ['first', 'odd', 'even']
               for b in ['Header', 'Footer']]
    position_names = ['left', 'center', 'right']

    return [ws.__getattribute__(h).__getattribute__(pos) for h in wsprops for pos in position_names]
