"""
This file is a simple example on how to use style sheets. If you want to use customized
font styles in your project, you should create a customized version.
@author: Nathanael Jöhrmann
"""
from pptx_tools.font_style import PPTXFontStyle
from pptx_tools.paragraph_style import PPTXParagraphStyle
# from pptx.enum.lang import MSO_LANGUAGE_ID
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT


# note: default values added to PPTXFontStyle as class attributes in v0.0.5
# define default values here, so you can easily change for whole presentation
# MY_DEFAULT_LANGUAGE = MSO_LANGUAGE_ID.ENGLISH_UK  # MSO_LANGUAGE_ID.GERMAN
# MY_DEFAULT_FONT_NAME = "Roboto"  # "Arial"  # "Arial Narrow"


def font_default() -> PPTXFontStyle:  # font for normal text
    result = PPTXFontStyle()
    # result.language_id = MY_DEFAULT_LANGUAGE
    # result.name = MY_DEFAULT_FONT_NAME
    result.size = 14
    return result


def font_small_text() -> PPTXFontStyle:
    result = font_default()
    result.size -= 2
    return result


def font_title() -> PPTXFontStyle:  # font for presentation title
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
    result.font = font_default()
    result.alignment = PP_PARAGRAPH_ALIGNMENT.LEFT
