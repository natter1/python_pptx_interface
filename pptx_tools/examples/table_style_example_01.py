"""
This script demonstrates how to work with tables and table-styles in python-pptx-interface.
@author: Nathanael Jöhrmann
"""
import os

from pptx_tools.creator import PPTXCreator
from pptx_tools.font_style import PPTXFontStyle
from pptx_tools.position import PPTXPosition
from pptx_tools.table_style import PPTXTableStyle
from pptx_tools.templates import TemplateExample


def run(save_dir: str):
    pp = PPTXCreator(TemplateExample())
    slide_01 = pp.add_slide("Table style example 01 - slide 01")
    slide_02 = pp.add_slide("Table style example 01 - slide 02")
    slide_03 = pp.add_slide("Table style example 01 - slide 03")
    slide_04 = pp.add_slide("Table style example 01 - slide 04")
    slide_05 = pp.add_slide("Table style example 01 - slide 05")
    slide_06 = pp.add_slide("Table style example 01 - slide 06")

    # data for a table with 5 rows and 3 cols.
    table_data = []
    table_data.append([1, "The second column is longer."])  # rows can have different length
    table_data.append([2, "Table entries don't have to be strings,"])  # there is specific type needed for entries (implemented as text=f"{entry}")
    table_data.append([3, "because its implemented as text=f'{entry}'"])
    table_data.append([4, "also note: the number of entries per row is", " not fixed"])
    table_data.append([5, "That's it for now."])

    # We can add these data as a table title_slide with PPTXCreator.add_table()...
    table_01 = pp.add_table(slide_01, table_data)
    # ... but if you open the slide in PowerPoint there are a few issues:
    #         1) the table is positioned in the top left corner - overlapping the title
    #         2) the first row is formated differently (like a column header)
    #         3) all columns have the same width (1 inch)
    #         4) the table width is too small (for the used font size)

    # Lets handle the the position first using optional PPTXPosition parameter
    table_02 = pp.add_table(slide_02, table_data, PPTXPosition(0.02, 0.14))

    # for more control we use PPTXTableStyle
    table_style = PPTXTableStyle()
    table_style.first_row_header = False
    table_style.width = 6.1  # table width in inches
    table_style.col_ratios = [0.3, 5, 1.3]

    table_03 = pp.add_table(slide_03, table_data, PPTXPosition(0.02, 0.14), table_style)

    # It's also possible to add the position directly to the table style:
    table_style.position = PPTXPosition(0.02, 0.14)
    # or to set the table width as a fraction of slide width:
    table_style.set_width_as_fraction(0.49)
    # change row/col bending
    table_style.col_banding = True
    table_style.row_banding = False
    table_04 = pp.add_table(slide_04, table_data, table_style=table_style)

    # we could also add a font-style and a cell-style
    table_style.font_style = PPTXFontStyle().set(italic=True, name="Arial", color_rgb=(100, 200, 30))
    # todo: cell-style
    table_05 = pp.add_table(slide_05, table_data, table_style=table_style)

    # you could also use a table style on an existing table
    table_06 = pp.add_table(slide_06, table_data)
    table_style.write_shape(table_06)

    pp.save(os.path.join(save_dir, "table_style_example_01.pptx"), overwrite=True)


if __name__ == '__main__':
    save_dir = os.path.dirname(os.path.abspath(__file__)) + '\\output\\'
    if not os.path.exists(save_dir):
        os.makedirs(save_dir)
    run(save_dir)
