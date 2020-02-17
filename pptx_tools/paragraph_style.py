"""
This module provides a helper class to deal with paragraphs in python-pptx.
@author: Nathanael Jöhrmann
"""
from typing import Optional

from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT
from pptx.text.text import _Paragraph

from pptx_tools.font_style import PPTXFontStyle


class PPTXParagraphStyle:
    """
    Helper class to deal with paragraphs in python-pptx.
    """
    def __init__(self):
        # the following values cannot be inherited from parent object(?) -----------------------------------------------
        self.alignment: Optional[PP_PARAGRAPH_ALIGNMENT] = None  # PP_PARAGRAPH_ALIGNMENT.CENTER/JUSTIFY/LEFT/RIGHT/...
        self.level: int = 0  # 0 .. 8 (indentation level)
        self.line_spacing: Optional[float] = None
        self.space_before: Optional[float] = None
        self.space_after: Optional[float] = None
        # --------------------------------------------------------------------------------------------------------------
        self.font: Optional[PPTXFontStyle] = None

    def write_paragraph(self, paragraph: _Paragraph) -> None:
        """Write paragraph style to given paragraph."""
        if self.alignment is not None:
            paragraph.alignment = self.alignment
        if self.level is not None:
            paragraph.level = self.level
        if self.line_spacing is not None:
            paragraph.line_spacing = self.line_spacing
        if self.font is not None:
            self.font.write_paragraph(paragraph)
