"""
This script demonstrates how to work with fonts and font-styles in python-pptx-interface.
@author: Nathanael Jöhrmann
"""

import os

from pptx.enum.lang import MSO_LANGUAGE_ID

from pptx_tools.creator import PPTXCreator
from pptx_tools.font_style import PPTXFontStyle
from pptx_tools.style_sheets import font_title
from pptx_tools.templates import TemplateExample


def run(save_dir: str):
    pp = PPTXCreator(TemplateExample())

    PPTXFontStyle.lanaguage_id = MSO_LANGUAGE_ID.ENGLISH_UK
    PPTXFontStyle.name = "Roboto"

    title_slide = pp.add_title_slide("Example presentation")
    font = font_title()  # returns a PPTXFontStyle instance with bold font and size = 32 Pt
    font.write_shape(title_slide.shapes.title)  # change font attributes for all paragraphs in shape


if __name__ == '__main__':
    save_dir = os.path.dirname(os.path.abspath(__file__)) + '\\output\\'
    if not os.path.exists(save_dir):
        os.makedirs(save_dir)
    run(save_dir)