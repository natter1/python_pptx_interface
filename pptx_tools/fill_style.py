"""
This module provides a helper class to deal with fills (for shapes, table cells ...) in python-pptx.
@author: Nathanael Jöhrmann
"""
from enum import Enum, auto

from pptx.dml.color import RGBColor


class FillType(Enum):
    NOFILL = auto()  # fill.background()
    SOLID = auto()  # fill.solid()
    PATTERNED = auto()  # # fill.paterned()
    GRADIENT = auto  # fill.gradient()


class PPTXFillStyle():
    def __init__(self):
        self.fill_type = FillType.SOLID
        self.fore_color = None
        self.fore_color = RGBColor(0xFB, 0x8F, 0x00)  # todo: test only
        self.fore_color_brightness = None

    def write_fill(self, fill: any):  # todo typing
        if self.fill_type is not None:
            self._write_fill_type(fill)

    def _write_fill_type(self, fill: any):  # todo: typing
        if self.fill_type == FillType.NOFILL:
            fill.background()

        elif self.fill_type == FillType.SOLID:
            if self.fore_color is not None:
                fill.solid()
                fill.fore_color = self.fore_color
                if self.fore_color_brightness:
                    fill.fore_color.brightness = self.fore_color_brightness
            else:
                print("Warning: Cannot set FillType.SOLID without a fore_color.")

        elif self.fill_type == FillType.PATTERNED:
            print("FillType.PATTERNED not implemented jet.")

        elif self.fill_type == FillType.GRADIENT:
            print("FillType.GRADIENT not implemented jet.")