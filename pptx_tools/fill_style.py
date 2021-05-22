"""
This module provides a helper class to deal with fills (for shapes, table cells ...) in python-pptx.
@author: Nathanael Jöhrmann
"""
from enum import Enum, auto
from typing import Union, Optional, Tuple

from pptx.dml.color import RGBColor
from pptx.dml.fill import FillFormat
from pptx.enum.base import EnumValue
from pptx.enum.dml import MSO_PATTERN_TYPE

from pptx_tools.utils import _DO_NOT_CHANGE


class FillType(Enum):
    NOFILL = auto()  # fill.background()
    SOLID = auto()  # fill.solid()
    PATTERNED = auto()  # # fill.patterned()
    GRADIENT = auto()  # fill.gradient(); not implemented jet


class PPTXFillStyle:
    def __init__(self):
        self.fill_type: Optional[FillType] = None  # FillType.SOLID
        self._fore_color_rgb: Union[RGBColor, Tuple[float, float, float], None] = None
        self._fore_color_mso_theme: Optional[EnumValue] = None
        self.fore_color_brightness: Optional[float] = None

        self._back_color_rgb: Union[RGBColor, Tuple[float, float, float], None] = None
        self._back_color_mso_theme: Optional[EnumValue] = None
        self.back_color_brightness: Optional[float] = None

        self.pattern: Optional[MSO_PATTERN_TYPE] = None  # 0 ... 47

    @property
    def fore_color_rgb(self) -> Optional[RGBColor]:
        return self._fore_color_rgb

    @property
    def fore_color_mso_theme(self) -> Optional[EnumValue]:
        return self._fore_color_mso_theme

    @property
    def back_color_rgb(self) -> Optional[RGBColor]:
        return self._back_color_rgb

    @property
    def back_color_mso_theme(self) -> Optional[EnumValue]:
        return self._back_color_mso_theme

    @fore_color_rgb.setter
    def fore_color_rgb(self, value: Union[RGBColor, Tuple[any, any, any], None]):
        if value is not None:
            assert isinstance(value, (RGBColor, tuple))
            self._fore_color_mso_theme = None  # only one color definition at a time!
        self._fore_color_rgb = RGBColor(*value) if isinstance(value, tuple) else value

    @fore_color_mso_theme.setter
    def fore_color_mso_theme(self, value: Optional[EnumValue]):
        if value is not None:
            assert isinstance(value, EnumValue)
            self._fore_color_rgb = None  # only one color definition at a time!
        self._fore_color_mso_theme = value

    @back_color_rgb.setter
    def back_color_rgb(self, value: Union[RGBColor, Tuple[any, any, any], None]):
        if value is not None:
            assert isinstance(value, (RGBColor, tuple))
            self._fore_color_mso_theme = None  # only one color definition at a time!
        self._back_color_rgb = RGBColor(*value) if isinstance(value, tuple) else value

    @back_color_mso_theme.setter
    def back_color_mso_theme(self, value: Optional[EnumValue]):
        if value is not None:
            assert isinstance(value, EnumValue)
            self._back_color_rgb = None  # only one color definition at a time!
        self._back_color_mso_theme = value

    def set(self, fill_type: FillType = _DO_NOT_CHANGE,
            fore_color_rgb: Union[RGBColor, Tuple[any, any, any], None] = _DO_NOT_CHANGE,
            fore_color_mso_theme: Optional[EnumValue] = _DO_NOT_CHANGE,
            fore_color_brightness: Optional[float] = _DO_NOT_CHANGE,
            back_color_rgb: Union[RGBColor, Tuple[any, any, any], None] = _DO_NOT_CHANGE,
            back_color_mso_theme: Optional[EnumValue] = _DO_NOT_CHANGE,
            back_color_brightness: Optional[float] = _DO_NOT_CHANGE,
            pattern: Optional[MSO_PATTERN_TYPE] = _DO_NOT_CHANGE
            ):
        """Convenience method to set several fill attributes together."""
        if fill_type is not _DO_NOT_CHANGE:
            self.fill_type = fill_type

        if fore_color_rgb is not _DO_NOT_CHANGE:
            self.fore_color_rgb = fore_color_rgb
        if fore_color_mso_theme is not _DO_NOT_CHANGE:
            self.fore_color_mso_theme = fore_color_mso_theme
        if fore_color_brightness is not _DO_NOT_CHANGE:
            self.fore_color_brightness = fore_color_brightness

        if back_color_rgb is not _DO_NOT_CHANGE:
            self.back_color_rgb = back_color_rgb
        if back_color_mso_theme is not _DO_NOT_CHANGE:
            self.back_color_mso_theme = back_color_mso_theme
        if back_color_brightness is not _DO_NOT_CHANGE:
            self.back_color_brightness = back_color_brightness

        if pattern is not _DO_NOT_CHANGE:
            self.pattern = pattern

    def write_fill(self, fill: FillFormat):
        """Write attributes to a FillFormat object."""
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

    def _write_back_color(self, fill: FillFormat):
        if self.back_color_rgb is not None:
            fill.back_color.rgb = self.back_color_rgb
        elif self.back_color_mso_theme is not None:
            fill.back_color.theme_color = self.back_color_mso_theme
        else:
            raise ValueError("No valid rgb_color set")
        if self.back_color_brightness:
            fill.back_color.brightness = self.back_color_brightness

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
            if self.pattern is not None:
                fill.pattern = self.pattern
            if (self.fore_color_rgb is not None) or (self.fore_color_mso_theme is not None):
                self._write_fore_color(fill)
            if (self.back_color_rgb is not None) or (self.back_color_mso_theme is not None):
                self._write_back_color(fill)

        elif self.fill_type == FillType.GRADIENT:
            print("FillType.GRADIENT not implemented jet.")
