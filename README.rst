python_pptx_interface
=====================
.. image:: https://img.shields.io/pypi/v/python_pptx_interface.svg
    :target: https://pypi.org/project/python_pptx_interface/

.. image:: http://img.shields.io/:license-MIT-blue.svg?style=flat-square
    :target: http://badges.MIT-license.org

`python-pptx <https://github.com/scanny/python-pptx.git>`_ is a great module to create pptx-files.
But it it can be challenging to master the complex syntax. This module tries to present an easier interface
for python-pptx to create simple PowerPoint files. It also adds some still missing features like moving slides,
create links to other slides or remove unused place-holders.

The main parts are:
  * PPTXCreator: Create pptx-File from template, incluing methods to add text, tables, figures etc.
  * PPTXFontStyle: Helps to set/change/copy fonts.
  * AbstractTemplate: Base class for all custom templates (enforce necessary attributes)
  * TemplateExample: Example class to show how to work with custom templates
  * utils.py - a collection of useful functions, eg. to generate PDF or PNG from \*.pptx (needs PowerPoint installed)


Example
-------

.. figure:: https://github.com/natter1/python_pptx_interface/raw/master/docs/images/example01_title_slide.png
    :width: 500pt

|

This module comes with an `example <https://github.com/natter1/python_pptx_interface/blob/master/pptx_tools/example.py>`_,
that you could run like

.. code:: python

    import pptx_tools.example as example

    example.run(save_dir=my_dir)  # you have to specify the folder where to save the presentation

This will create an example.pptx, using some of the key-features of python-pptx-interface. Lets have a closer look:

|

.. code:: python

    from pptx_tools.creator import PPTXCreator, PPTXPosition
    from pptx_tools.style_sheets import font_title, font_default
    from pptx_tools.templates import TemplateExample

    from pptx.enum.lang import MSO_LANGUAGE_ID
    from pptx.enum.text import MSO_TEXT_UNDERLINE_TYPE

    try:
        import matplotlib.pyplot as plt
        matplotlib_installed = True
    except ImportError as e:
        matplotlib_installed = False

First we need to import some stuff. **PPTXCreator** is the class used to create a \*.pptx file.
**PPTXPosition** allows as to position shapes in more intuitive units of slide width/height.
**font_title** is a function returning a PPTXFontStyle instance. We will use it to change the formatting of the title shape.
**TemplateExample** is a class providing access to the example-template.pptx included in python-pptx-interface
and also setting some texts on the master slides like author, date and website. You could use it as reference
on how to use your own template files by subclassing AbstractTemplate
(you need at least to specify a path to your template and define a default_layout and a title_layout).

**MSO_LANGUAGE_ID** is used to set the language of text and **MSO_TEXT_UNDERLINE_TYPE** is used to format underlining.

Importing matplotlib is optional - it is used to demonstrate, how to get a matplotlib figure into your presentation.

|
|

.. code:: python

    def run(save_dir: str):
        pp = PPTXCreator(TemplateExample())

        PPTXFontStyle.lanaguage_id = MSO_LANGUAGE_ID.ENGLISH_UK
        PPTXFontStyle.name = "Roboto"

        title_slide = pp.add_title_slide("Example presentation")
        font = font_title()  # returns a PPTXFontStyle instance with bold font and size = 32 Pt
        font.write_shape(title_slide.shapes.title)  # change font attributes for all paragraphs in shape

Now we create our presentation with **PPTXCreator** using the **TemplateExample**.
We also set the default font language and name of all **PPTXFontStyle** instances. This is not necessary,
as *ENGLISH_UK* and *Roboto* are the defaults anyway. But in principle you could change these settings here,
to fit your needs. If you create your own template class, you might also set these default parameters there.
Finally we add a title slide and change the font style of the title using title_font().

|
|

.. code:: python

        slide2 = pp.add_slide("page2")
        pp.add_slide("page3")
        pp.add_slide("page4")
        content_slide = pp.add_content_slide()  # add slide with hyperlinks to all other slides

Next, we add three slides, and create a content slide with hyperlinks to all other slides. By default,
it is put to the second position (you could specify the position using the optional slide_index parameter).

.. figure:: https://github.com/natter1/python_pptx_interface/raw/master/docs/images/example01_content_slide.png
    :width: 500pt

|
|

Lets add some more stuff to the title slide.

.. code:: python

    text = "This text has three paragraphs. This is the first.\n" \
           "Das ist der zweite ...\n" \
           "... and the third."
    my_font = font_default()  # font size 14
    my_font.size = 16
    text_shape = pp.add_text_box(title_slide, text, PPTXPosition(0.02, 0.24), my_font)

**PPTXCreator.add_text_box()** places a new text shape on a slide with the given text.
Optionally it accepts a PPTXPosition and a PPTXFont. With PPTXPosition(0.02, 0.24)
we position the figure 0.02 slide widths from left and 0.24 slide heights from top.

|
|

.. code:: python

    my_font.set(size=22, bold=True, language_id=MSO_LANGUAGE_ID.GERMAN)
    my_font.write_paragraph(text_shape.text_frame.paragraphs[1])

    my_font.set(size=18, bold=False, italic=True, name="Vivaldi",
                language_id=MSO_LANGUAGE_ID.ENGLISH_UK,
                underline=MSO_TEXT_UNDERLINE_TYPE.WAVY_DOUBLE_LINE)
    my_font.write_paragraph(text_shape.text_frame.paragraphs[2])

We can use my_font to format individual paragraphs in a text_frame with **PPTXFontStyle.write_paragraph()**.
Via **PPTXFontStyle.set()** easily customize the font before using it.

|
|

.. code:: python

        table_data = []
        table_data.append([1, 2])  # rows can have different length
        table_data.append([4, slide2, 6])  # there is specific type needed for entries (implemented as text=f"{entry}")
        table_data.append(["", 8, 9])

        pp.add_table(slide2, table_data)

we can also easily add a table. First we define all the data we want to put in the table. Here we use a list of lists.
But add_table is more flexible and can work with anything, that is an Iterable of Iterable. The outer iterable defines,
how many rows the table will have. The longest inner iterable is used to get the number of columns.

|
|

.. code:: python

        if matplotlib_installed:
            fig = create_demo_figure()
            pp.add_matplotlib_figure(fig, title_slide, PPTXPosition(0.3, 0.4))
            pp.add_matplotlib_figure(fig, title_slide, PPTXPosition(0.3, 0.4, fig.get_figwidth(), -1.0), zoom=0.4)
            pp.add_matplotlib_figure(fig, title_slide, PPTXPosition(0.3, 0.4, fig.get_figwidth(), 0.0), zoom=0.5)
            pp.add_matplotlib_figure(fig, title_slide, PPTXPosition(0.3, 0.4, fig.get_figwidth(), 1.5), zoom=0.6)


If matplotlib is installed, we use it to create a demo figure, and add it to the title_slide.
With PPTXPosition(0.3, 0.4) we position the figure 0.3 slide widths from left and 0.4 slide heights from top.
PPTXPosition has two more optional parameters, to further position with inches values (starting from the relative position).

|
|

.. code:: python

        pp.save(os.path.join(save_dir, "example.pptx"))

Finally, we save the example as example.pptx.

|
|

If you are on windows an have PowerPoint installed, you could use some additional features.

.. code:: python

    try:  # only on Windows with PowerPoint installed:
        filename_pptx = os.path.join(save_dir, "example.pptx")
        filename_pdf = os.path.join(save_dir, "example.pdf")
        foldername_png = os.path.join(save_dir, "example_pngs")

        # use absolute path, because its not clear where PowerPoint saves PDF/PNG ... otherwise
        pp.save(filename_pptx, create_pdf=True)  # saves your pptx-file and also creates a PDF file
        pp.save_as_pdf(filename_pdf, overwrite=True)  # saves presentation as PDF
        pp.save_as_png(foldername_png, overwrite_folder=True)  # creates folder with PNGs of slides
    except Exception as e:
        print(e)


Requirements
------------

* Python >= 3.6 (f-strings)
* python-pptx

Optional requirements
---------------------
* matplotlib (adding matplotlib figures to presentation)
* comtypes  (create PDF's or PNG's)
* PowerPoint (create PDF's or PNG's)

Contribution
------------
Help with this project is welcome. You could report bugs or ask for improvements by creating a new issue.

If you want to contribute code, here are some additional notes:

* This project uses 120 characters per line.
* Try to avoid abbreviations in names for functions or variables.
* Use type hints.
* Use Slide objects instead of IDs or index values as function parameter.
