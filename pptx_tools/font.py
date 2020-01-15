"""
This module provides a helper class to deal with fonts in python-pptx.
@author: Nathanael Jöhrmann
"""
from typing import Union

from pptx.enum.lang import MSO_LANGUAGE_ID
from pptx.util import Pt

from pptx.enum.text import MSO_TEXT_UNDERLINE_TYPE
import pptx.text.text.Font


class PPTXFontTool:
    """
    Helper class to deal with fonts in python-pptx. The internal class pptx.text.text.Font is limited, as it
    always needs an existing Text/Character/... for initializing and also basic functionality like assignment
    of one font to another is missing.
    """

    def __init__(self):
        #  If set to None, the bold and italic ... setting is cleared and is inherited
        #  from an enclosing shape’s setting, or a setting in a style or master
        self.bold: Union[bool, None] = None
        self.italic: Union[bool, None] = None
        self.language_id: MSO_LANGUAGE_ID = MSO_LANGUAGE_ID.NONE  # ENGLISH_UK; ENGLISH_US; ESTONIAN; GERMAN; ...
        self.name: Union[str, None] = None

        # saved in units of Pt (not EMU like pptx.text.text.Font) - convertation to EMU is done during write_to_font
        self.size: Union[int, None] = None  # 18
        self.underline: Union[MSO_TEXT_UNDERLINE_TYPE, bool, None] = None

        # todo: color is ColorFormat object
        # todo: fill is FillFormat object
        # self.color = ...
        # self.fill = ...

    def read_from_font(self, font: pptx.text.text.Font):
        font.bold = self.bold
        font.italic = self.italic
        font.language_id = self.language_id
        font.name = self.name
        font.size = Pt(self.size)
        font.underline = self.underline

    def write_to_font(self, font: pptx.text.text.Font):
        font.bold = self.bold
        font.italic = self.italic
        font.language_id = self.language_id
        font.name = self.name
        font.size = Pt(self.size)
        font.underline = self.underline

    @classmethod
    def copy_font(cls, _from, _to):
        _to.bold = _from.bold
        # todo: color is ColorFormat object
        # _to.set_color = _from.color
        # todo: fill is FillFormat object
        # _to.fill = _from.fill
        _to.italic = _from.italic
        _to.language_id = _from.language_id
        _to.name = _from.name
        _to.size = _from.size
        _to.underline = _from.underline
