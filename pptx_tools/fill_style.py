"""
This module provides a helper class to deal with fills (for shapes, table cells ...) in python-pptx.
@author: Nathanael Jöhrmann
"""
from enum import Enum, auto
from typing import Union, Optional, Tuple

from pptx.dml.color import RGBColor
from pptx.dml.fill import FillFormat
from pptx.enum.base import EnumValue
from pptx.enum.dml import MSO_THEME_COLOR_INDEX


class FillType(Enum):
    NOFILL = auto()  # fill.background()
    SOLID = auto()  # fill.solid()
    PATTERNED = auto()  # # fill.patterned()
    GRADIENT = auto  # fill.gradient()


class PPTXFillStyle:
    def __init__(self):
        self.fill_type = FillType.SOLID
        self._fore_color_rgb = None
        self._fore_color_mso_theme = None
        self.fore_color_brightness = None

    @property
    def fore_color_rgb(self):
        return self._fore_color_rgb

    @fore_color_rgb.setter
    def fore_color_rgb(self, value: Union[RGBColor, Tuple[any, any, any], None]):
        if value is not None:
            assert isinstance(value, RGBColor) or isinstance(value, tuple)
            self._fore_color_mso_theme = None  # only one color definition at a time!
        if isinstance(value, tuple):
            self._fore_color_rgb = RGBColor(*value)
        else:
            self._fore_color_rgb = value

    @property
    def fore_color_mso_theme(self):
        return self._fore_color_mso_theme

    @fore_color_mso_theme.setter
    def fore_color_mso_theme(self, value):
        if value is not None:
            assert isinstance(value, EnumValue)
            self._fore_color_rgb = None  # only one color definition at a time!
        self._fore_color_mso_theme = value

    def write_fill(self, fill: FillFormat):
        if self.fill_type is not None:
            self._write_fill_type(fill)

    def _write_fore_color(self, fill: FillFormat):
        if self.fore_color_rgb is not None:
            fill.fore_color.rgb = self.fore_color_rgb
        elif self.fore_color_mso_theme is not None:
            fill.fore_color.theme_color = self.fore_color_mso_theme
        else:
            raise ValueError("No valid rgb_color set")
        if self.fore_color_brightness:
            fill.fore_color.brightness = self.fore_color_brightness

    def _write_fill_type(self, fill: FillFormat):
        if self.fill_type == FillType.NOFILL:
            fill.background()

        elif self.fill_type == FillType.SOLID:
            if (self.fore_color_rgb is not None) or (self.fore_color_mso_theme is not None):
                fill.solid()
                self._write_fore_color(fill)
            else:
                print("Warning: Cannot set FillType.SOLID without a valid fore_color_*.")

        elif self.fill_type == FillType.PATTERNED:
            fill.patterned()
            if (self.fore_color_rgb is not None) or (self.fore_color_mso_theme is not None):
                self._write_fore_color(fill)

        elif self.fill_type == FillType.GRADIENT:
            print("FillType.GRADIENT not implemented jet.")
