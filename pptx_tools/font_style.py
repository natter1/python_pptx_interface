"""
This module provides a helper class to deal with fonts in python-pptx.
@author: Nathanael Jöhrmann
"""
from typing import Union, Optional

from pptx.enum.lang import MSO_LANGUAGE_ID
from pptx.enum.text import MSO_TEXT_UNDERLINE_TYPE
from pptx.shapes.autoshape import Shape
from pptx.text.text import Font
from pptx.util import Pt
from pptx.text.text import _Run
from pptx.text.text import _Paragraph

class PPTXFontStyle:
    """
    Helper class to deal with fonts in python-pptx. The internal class pptx.text.text.Font is limited, as it
    always needs an existing Text/Character/... for initializing and also basic functionality like assignment
    of one font to another is missing.
    """
    # default language anf font
    language_id: MSO_LANGUAGE_ID = MSO_LANGUAGE_ID.ENGLISH_UK  # MSO_LANGUAGE_ID.GERMAN
    name = "Roboto"  # "Arial"  # "Arial Narrow"

    def __init__(self):
        #  If set to None, the bold and italic ... setting is cleared and is inherited
        #  from an enclosing shape’s setting, or a setting in a style or master
        self.bold: Optional[bool] = None
        self.italic: Optional[bool] = None

        # use class attribute; instance attribute only when changed by user
        # self.language_id: MSO_LANGUAGE_ID = MSO_LANGUAGE_ID.NONE  # ENGLISH_UK; ENGLISH_US; ESTONIAN; GERMAN; ...
        # self.name: Optional[str] = None

        # saved in units of Pt (not EMU like pptx.text.text.Font) - converting to EMU is done during write_to_font
        self.size: Optional[int] = None  # 18
        self.underline: Union[MSO_TEXT_UNDERLINE_TYPE, bool, None] = None

        # todo: color is ColorFormat object
        # todo: fill is FillFormat object
        # self.color = ...
        # self.fill = ...

    def read_font(self, font: Font) -> None:
        """Read attributes from a pptx.text.text.Font object."""
        font.bold = self.bold
        font.italic = self.italic
        font.language_id = self.language_id
        font.name = self.name
        if self.size is not None:
            font.size = Pt(self.size)
        font.underline = self.underline

    def write_font(self, font: Font) -> None:
        """Write attributes to a pptx.text.text.Font object."""
        font.bold = self.bold
        font.italic = self.italic
        font.language_id = self.language_id
        font.name = self.name
        font.size = Pt(self.size)
        font.underline = self.underline

    def write_shape(self, shape: Shape) -> None:  # todo: remove? better use write_text_fame
        """
        Write attributes to all paragraphs in given pptx.shapes.autoshape.Shape.
        Raises TypeError if given shape has no text_frame.
        """
        if not shape.has_text_frame:
            raise TypeError("Cannot write font for given shape (has no text_frame)")
        self.write_text_frame(shape.text_frame)

    def write_text_frame(self, text_frame):
        for paragraph in text_frame.paragraphs:
            self.write_paragraph(paragraph)

    def write_paragraph(self, paragraph: _Paragraph) -> None:
        """ Write attributes to given paragraph"""
        self.write_font(paragraph.font)

    def write_run(self, run: _Run) -> None:
        """ Write attributes to given run"""
        self.write_font(run.font)

    @classmethod
    def copy_font(cls, _from: Font, _to: Font) -> None:
        """Copies settings from one pptx.text.text.Font to another."""
        font_style=cls()
        font_style.read_font(_from)
        font_style.write_font(_to)
        # _to.bold = _from.bold
        # # todo: color is ColorFormat object
        # # _to.set_color = _from.color
        # # todo: fill is FillFormat object
        # # _to.fill = _from.fill
        # _to.italic = _from.italic
        # _to.language_id = _from.language_id
        # _to.name = _from.name
        # _to.size = _from.size
        # _to.underline = _from.underline

    def set(self, bold: Optional[bool] = None,
            italic: Optional[bool] = None,
            language_id: MSO_LANGUAGE_ID = None,
            name: Optional[str] = None,
            size: Optional[int] = None,
            underline: Union[MSO_TEXT_UNDERLINE_TYPE, bool, None] = None
            ):
        """Convienience method to set several font attributes together."""
        if bold is not None:
            self.bold = bold
        if italic is not None:
            self.italic = italic
        if language_id is not None:
            self.language_id = language_id
        if name is not None:
            self.name = name
        if size is not None:
            self.size = size
        if underline is not None:
            self.underline = underline

        # -----------------------------------------------------------------------------------------------
        # ----------------------------------- experimentell methods -------------------------------------
        # -----------------------------------------------------------------------------------------------
    def _write_font_experimentell(self, font: Font,all_caps: bool = True, strikethrough: bool = True):
        if all_caps:
            font._element.attrib['cap'] = "all"
        else:
            pass

        if strikethrough:
            font._element.attrib['strike'] = "sngStrike"
        else:
            pass
