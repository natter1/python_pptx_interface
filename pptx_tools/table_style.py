"""
This module provides a helper class to deal with tables in python-pptx.
@author: Nathanael Jöhrmann
"""
from enum import Enum, auto
from typing import Generator

from pptx.dml.color import RGBColor
from pptx.shapes.autoshape import Shape
from pptx.table import Table, _Cell

from pptx_tools.fill_style import PPTXFillStyle
from pptx_tools.font_style import PPTXFontStyle


class PPTXCellStyle:  # format tale cell
    def __init__(self):
        self.fill_style = PPTXFillStyle()

    def write_cell(self, cell: _Cell) -> None:
        self.fill_style.write_fill(cell.fill)


class PPTXTableStyle:
    """
    ...
    """
    def __init__(self):
        self.font_style = PPTXFontStyle()
        self.cell_style = PPTXCellStyle()
        self.first_row_header = None  # False  # special formatting for first row?
        self.band_col = None  # False  # slightly alternate color brightness per col
        self.band_row = None  #True  # slightly alternate color brightness per row

        self.width = None
        self.cols_ratio = None

    def iter_cells(self, table: Table) -> Generator[_Cell, None, None]:
        for row in table.rows:
            for cell in row.cells:
                yield cell

    def write_shape(self, shape: Shape) -> None:
        if not shape.has_table:
            print(f"Warning: Could not write table style. {shape} has no table.")
            return
        table: Table = shape.table

        if self.first_row_header is not None:
            table._tbl.firstRow = self.first_row_header

        if self.band_col is not None:
            table._tbl.band_col = self.band_col

        if self.band_row is not None:
            table._tbl.band_row = self.band_row

        # font is managed per cell; there is no "table font"
        for cell in self.iter_cells(table):
            self.font_style.write_text_frame(cell.text_frame)



