python_pptx_interface
=====================

Easy interface to create pptx-files using python-pptx.
  * PPTXCreator: Create pptx-File from template, incluing methods to add text, figures etc.
  * PPTXFontTool: Helps to set/change/copy fonts.
  * TemplateExample: Example class to show how to work with custom templates
  * AbstractTemplate: Base class for all custom templates (enforce necessary attributes)


Contribution
------------
Help with this project is welcome. You could report bugs or ask for improvements by creating a new issue.

If you want to contribute code, here are some additional notes:

* This project uses 120 characters per line.
* Try to avoid abbreviations in names for functions or variables.
* Use type hints.
* Use Slide objects instead of IDs or index values as function parameter.


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


