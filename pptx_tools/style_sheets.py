"""
This file is a simple example on how to use style sheets. If you want to use customized
paragraph styles in your project, you should create a customized version.
@author: Nathanael Jöhrmann
"""

# from pptx.enum.lang import MSO_LANGUAGE_ID
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT

from pptx_tools.fill_style import FillType
from pptx_tools.font_style import PPTXFontStyle
from pptx_tools.paragraph_style import PPTXParagraphStyle
from pptx_tools.table_style import PPTXTableStyle, PPTXCellStyle


# note: default values added to PPTXFontStyle as class attributes in v0.0.5
# MY_DEFAULT_LANGUAGE = MSO_LANGUAGE_ID.ENGLISH_UK  # MSO_LANGUAGE_ID.GERMAN
# MY_DEFAULT_FONT_NAME = "Roboto"  # "Arial"  # "Arial Narrow"


# table_style.cell_style.fill_style.fore_color_mso_theme = MSO_THEME_COLOR_INDEX.ACCENT_1
# table_style.cell_style.fill_style.fore_color_rgb = (200, 200, 200)
# table_style.font_style.size = 10
# table_style.font_style.bold = True

def table_invisible() -> PPTXTableStyle:
    result = PPTXTableStyle()
    result.cell_style = PPTXCellStyle()
    result.cell_style.fill_style.fill_type = FillType.NOFILL
    # todo: implement control for border lines
    return result


def table_no_header() -> PPTXTableStyle:
    result = PPTXTableStyle()

    result.first_row_header = False
    result.row_banding = True
    result.col_banding = False

    return result


def font_default() -> PPTXFontStyle:  # paragraph for normal text
    result = PPTXFontStyle()
    # result.language_id = MY_DEFAULT_LANGUAGE
    # result.name = MY_DEFAULT_FONT_NAME
    result.size = 14
    return result


def font_small_text() -> PPTXFontStyle:
    result = font_default()
    result.size -= 2
    return result


def font_title() -> PPTXFontStyle:  # paragraph for presentation title
    result = font_default()
    result.size = 32
    result.bold = True
    return result


def font_slide_title() -> PPTXFontStyle:
    result = font_title()
    result.size = 28
    return result


def font_sub_title() -> PPTXFontStyle:
    result = font_title()
    result.size = 18
    return result


def paragraph_default():
    result = PPTXParagraphStyle()
    result.font_style = font_default()
    result.alignment = PP_PARAGRAPH_ALIGNMENT.LEFT
