python_pptx_interface
=====================
.. image:: https://img.shields.io/pypi/v/python_pptx_interface.svg
    :target: https://pypi.org/project/python_pptx_interface/

.. image:: http://img.shields.io/:license-MIT-blue.svg?style=flat-square
    :target: http://badges.MIT-license.org

`python-pptx <https://github.com/scanny/python-pptx.git>`_ is a great module to create pptx-files.
But it it can be challenging to master the complex syntax. This module tries to present an easier interface
for python-pptx to create simple PowerPoint files. It also add some still missing features like moving slides,
create links to other slides or remove unused place-holders.

The main parts are:
  * PPTXCreator: Create pptx-File from template, incluing methods to add text, tables, figures etc.
  * PPTXFontTool: Helps to set/change/copy fonts.
  * AbstractTemplate: Base class for all custom templates (enforce necessary attributes)
  * TemplateExample: Example class to show how to work with custom templates

Requirements
------------

* Python >= 3.6 (f-strings)
* python-pptx

Contribution
------------
Help with this project is welcome. You could report bugs or ask for improvements by creating a new issue.

If you want to contribute code, here are some additional notes:

* This project uses 120 characters per line.
* Try to avoid abbreviations in names for functions or variables.
* Use type hints.
* Use Slide objects instead of IDs or index values as function parameter.

Example
-------
This module comes with an `example <https://github.com/natter1/python_pptx_interface/blob/master/pptx_tools/example.py>`_,
that you could run like

.. code:: python

    import pptx_tools.example as example

    example.run()

This will create an example.pptx, using some of the key-features of python-pptx-interface. Lets have a closer look:

.. code:: python

    from pptx_tools.creator import PPTXCreator, PPTXPosition
    from pptx_tools.style_sheets import font_title
    from pptx_tools.templates import TemplateExample

    try:
        import matplotlib.pyplot as plt
        matplotlib_installed = True
    except ImportError as e:
        matplotlib_installed = False

First we need to import some stuff. **PPTXCreator** is the class used to create a \*.pptx file.
**PPTXPosition** allows as to position shapes in more intuitive units of slide width/height.
**font_title** is a function returning a FontStyleTool instance. We will use it to change the formatting of the title shape.
**TemplateExample** is a class providing access to the example-template.pptx included in python-pptx-interface
and also setting some texts on the master slides like author, date and website. You could use it as reference
on how to use your own template files by subclassing AbstractTemplate
(you need at least to specify a path to your template and define a default_layout and a title_layout).

Importing matplotlib is optional - it is used to demonstrate, how to get a matplotlib figure into your presentation.

.. code:: python

    def run():
        pp = PPTXCreator(TemplateExample())

        title_slide = pp.add_title_slide("Example presentation")
        font = font_title()
        font.write_shape(title_slide.shapes.title)  # you can change font attributes of paragraphs in shape via PPTXFontTool

Now we create our presentation, add a title slide and change the font style of the title using title_font().

.. code:: python

        slide2 = pp.add_slide("page2")
        pp.add_slide("page3")
        pp.add_slide("page4")
        content_slide = pp.add_content_slide()

Next, we add thre more slides, and create a content slide with hyperlinks to all other slides. By default,
it is put to the second position (you could specify the position using the optional slide_index parameter).

.. code:: python

        if matplotlib_installed:
            fig = create_demo_figure()
            pp.add_matplotlib_figure(fig, title_slide, PPTXPosition(0.3, 0.4))
            pp.add_matplotlib_figure(fig, title_slide, PPTXPosition(0.7, 0.4), zoom=0.4)

If matplotlib is installed, we use it to create a demo figure, and add it to the title_slide.
With PPTXPosition(0.3, 0.4) we position the figure 0.3 slide widths from left and 0.4 slide heights from top.
PPTXPosition has two more optional parameters, to further position with inches values (starting from the relative position).

.. code:: python

        table_data = []
        table_data.append([1, 2])  # rows can have different length
        table_data.append([4, slide2, 6])  # there is specific type needed for entries (implemented as text=f"{entry}")
        table_data.append(["", 8, 9])

        pp.add_table(slide2, table_data)

we can also easily add a table. First we define all the data we want to put in the table. Here we use a list of lists.
But add_table is more flexible and can work ith anything, thats an Iterable of Iterable. The outer iterable defines,
how many rows the table will have. The longest inner iterable is used to get the number of columns.

.. code:: python

        pp.save("example.pptx")

Finally, we save the example as example.pptx.

API
---

creator.py
...........

class PPTXCreator
  PPTXCreator(template: Union[Type[pptx_tools.templates.AbstractTemplate], NoneType] = None)

This Class provides an easy interface to create a PowerPoint presentation.
    - PPTXPosion is used to position new shapes (allowing position as fraction of slide height/width)
    - use pptx templates (in combination with templates.py)
    - removes unused placeholder from added slides

Methods defined here:

add_content_slide(self, slide_index=1)
    Adds a content slide with hyperlinks to all other slides and puts it to position slide_index.

  add_matplotlib_figure(self, fig: matplotlib.figure.Figure, slide: pptx.slide.Slide, pptx_position: pptx_tools.creator.PPTXPosition = None, zoom: float = 1.0, \*\*kwargs) -> pptx.shapes.picture.Picture
    Add a motplotlib figure to slide and position it via pptx_position.
    Optional parameter zoom sets image scaling in PowerPoint; only used if width not in kwargs (default = 1.0)

add_slide(self, title: str, layout=None) -> pptx.slide.Slide
    Adds a new slide to presentation. If now layout is given, default_layout is used.

add_text_box(self, slide, text: str, position: pptx_tools.creator.PPTXPosition = None, font: pptx_tools.font_style.PPTXFontStyle = None) -> pptx.shapes.autoshape.Shape
    Adds a text box with given text using given position and font.
    Uses self.default_position if no position is given.

add_title_slide(self, title: str, layout: pptx.slide.SlideLayout = None) -> pptx.slide.Slide
    Adds a new slide to presentation. If now layout is given, title_layout is used.

move_slide(self, slide: pptx.slide.Slide, new_index: int)
    Moves the given slide to position new_index.

save(self, filename: str) -> None
    Saves the presentation under the given filename.

Static methods defined here:

create_hyperlink(run: pptx.text.text._Run, shape: pptx.shapes.autoshape.Shape, to_slide: pptx.slide.Slide)
    Makes the given run a hyperlink to to_slide.

remove_unpopulated_shapes(slide: pptx.slide.Slide)
    Removes empty placeholders (e.g. due to layout) from slide.
    Further testing needed.


