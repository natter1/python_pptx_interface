"""
This module provides a helper class to deal with paragraphs in python-pptx.
@author: Nathanael Jöhrmann
"""
from typing import Union, Optional

from pptx.enum.lang import MSO_LANGUAGE_ID

from pptx.shapes.autoshape import Shape


from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT

from pptx_tools.font_style import PPTXFontStyle


class PPTXParagraphStyle:
    """
    Helper class to deal with paragraphs in python-pptx.
    """
    def __init__(self):
        self.alignment: Optional[PP_PARAGRAPH_ALIGNMENT] = None  # PP_PARAGRAPH_ALIGNMENT.CENTER/JUSTIFY/LEFT/RIGHT/...
        self.level: int = 0  # 0 .. 8 (indentation level)
        self.line_spacing: Optional[float] = None
        self.font: Optional[PPTXFontStyle] = None


