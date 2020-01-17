"""
This module provides an easier Interface to create *.pptx presentations using the module python-pptx.
@author: Nathanael Jöhrmann
"""
import io
from typing import Type, Optional

from matplotlib.figure import Figure
from pptx.presentation import Presentation
from pptx.shapes.autoshape import Shape
from pptx.shapes.picture import Picture
from pptx.text.text import _Run
from pptx.util import Inches
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.slide import Slide, SlideLayout

from pptx_tools.font_style import PPTXFontStyle
from pptx_tools.templates import AbstractTemplate


class PPTXPosition:
    """
    Used to generate positions of elements in slide coordiinates
    saves pptx presentation in class (PPTXPosition.prs) - this allows to call methods without setting prs each time:
    PPTXPosition(0.7, 0.3).dict()  # possible if instance was once initialized with valid presentation
    Creating a PPTXCreator instance is enough to set PPTXPosition.prs.
    when initializing a ) or PPTXPosition.prs was set directly.
    If you want to use several presentations with differing slide sizes in one script,
    you can set prs manually as an instance attribute:
    pptx1_position = PPTXPosition()  # uses PPTXPosition.prs
    pptx2_position = PPTXPosition()  # still uses PPTXPosition.prs, same as pptx1_position
    pptx2_position.prs = pptx2  # now uses pptx2
    """
    prs = None

    def __init__(self, left_rel=0.0, top_rel=0.0, left=0, top=0, presentation=None):
        """
        :param presentation: pptx.prs (needed for slide width and height)
        :param left_rel: distance from slide left (relative to slide width)
        :param top_rel: distance from slide top (relative to slide height)
        :param left: "left" to default_position figure [inches] starting from rel_left
        :param top: "top" to default_position figure [inches] starting from rel_top
        """

        if presentation:
            PPTXPosition.prs = presentation
        if not PPTXPosition.prs:
            raise Exception("When creating a PPTXPosition instance for the first time,"
                            " you have to provide a valid presentation")

        self.left_rel = left_rel
        self.top_rel = top_rel
        self.left = left
        self.top = top

    def set(self, left_rel=0.0, top_rel=0.0, left=0, top=0):
        self.left_rel = left_rel
        self.top_rel = top_rel
        self.left = left
        self.top = top

    def dict_for_position(self, left_rel=0.0, top_rel=0.0, left=0, top=0):
        """
        Returns kwargs dict for given default_position. Does not change attributes of self
        :param left_rel: float [slide_width]
        :param top_rel: float [slide_height]
        :param left: float [inch]
        :param top: float [inch]
        :return: dictionary
        """
        left = self.fraction_width_to_inch(left_rel) + left
        top = self.fraction_height_to_inch(top_rel) + top
        return {"left": left, "top": top}

    def dict(self):
        """
        This method returns a kwargs dict containing "left" and "top".
        :return: dictionary
        """
        return self.dict_for_position(self.left_rel, self.top_rel, self.left, self.top)

    def fraction_width_to_inch(self, fraction):
        """
        Returns a width in inches calculated as a fraction of total slide-width.
        :param fraction: float
        :return: Calculated Width in inch
        """
        result = Inches(self.prs.slide_width.inches * fraction)
        return result

    def fraction_height_to_inch(self, fraction):
        """
        Returns a height in inches calculated as a fraction of total slide-height.
        :param fraction: float
        :return: Calculated Width in inch
        """
        return Inches(self.prs.slide_height.inches * fraction)


# todo: template_file to Enum in pptx_template with all available templates
class PPTXCreator:
    """
    This Class provides an easy interface to create a PowerPoint presentation.
        - PPTXPosion is used to position new shapes (allowing position as fraction of slide height/width)
        - use pptx templates (in combination with templates.py)
        - removes unused placeholder from added slides
    """
    # disable typechecker, because None values are not allowed for attributes, but needed to create them in __init__
    # Correct values are set when calling self._create_presentation()
    # noinspection PyTypeChecker
    def __init__(self, template: Optional[Type[AbstractTemplate]] = None):
        self.slides: list = []
        self.template: Type[AbstractTemplate] = None
        self.prs: Presentation = None
        self.title_layout: SlideLayout = None
        self.default_layout: SlideLayout = None
        self._create_presentation(template)
        self.default_position = PPTXPosition(presentation=self.prs)

    def _fraction_width_to_inch(self, fraction: float) -> Inches:
        """
        Returns a width in inches calculated as a fraction of total slide-width.
        """
        result = Inches(self.prs.slide_width.inches * fraction)
        return result

    def _fraction_height_to_inch(self, fraction: float) -> Inches:
        """
        Returns a height in inches calculated as a fraction of total slide-height.
        """
        return Inches(self.prs.slide_height.inches * fraction)

    def save(self, filename: str) -> None:
        """
        Saves the presentation under the given filename.
        """
        self.prs.save(filename)

    def _create_presentation(self, template=None) -> None:
        """
        Create a new presentation (using optional template)
        """
        if template:
            self._create_presentation_from_template(template)
        else:
            self.prs = Presentation()
            self.title_layout = self.prs.slide_masters[0].slide_layouts[0]
            self.default_layout = self.prs.slide_masters[0].slide_layouts[0]

    def _create_presentation_from_template(self, template: Type[AbstractTemplate]) -> None:
        self.template = template
        self.prs = template.prs
        self.title_layout = template.title_layout
        self.default_layout = template.default_layout

    def add_title_slide(self, title: str, layout: SlideLayout = None) -> Slide:
        """Adds a new slide to presentation. If now layout is given, title_layout is used."""
        if not layout:
            layout = self.title_layout
        return self.add_slide(title, layout)

    def add_slide(self, title: str, layout=None) -> Slide:
        """Adds a new slide to presentation. If now layout is given, default_layout is used."""
        if not layout:
            layout = self.default_layout
        slide = self.prs.slides.add_slide(layout)
        title_shape = slide.shapes.title
        title_shape.text = title
        self.remove_unpopulated_shapes(slide)
        return slide

    def _write_position_in_kwargs(self, kwargs: dict, left_rel: float = 0.0, top_rel: float = 0.0) -> None:
        """
        This method modifies(!) the argument kwargs by adding or changing the entries "left" and "top".
        """
        left = self._fraction_width_to_inch(left_rel)
        top = self._fraction_height_to_inch(top_rel)

        if not "left" in kwargs:
            kwargs["left"] = left
        else:
            kwargs["left"] = kwargs["left"] + left
        if not "top" in kwargs:
            kwargs["top"] = top
        else:
            kwargs["top"] = kwargs["top"] + top

    def add_matplotlib_figure(self, fig: Figure, slide: Slide,
                              pptx_position: PPTXPosition = None,
                              zoom: float = 1.0,
                              **kwargs) -> Picture:
        """
        Add a motplotlib figure to slide and position it via pptx_position.
        Optional parameter zoom sets image scaling in PowerPoint; only used if width not in kwargs (default = 1.0)
        """
        if "width" not in kwargs:
            kwargs["width"] = Inches(fig.get_figwidth() * zoom)
        if not pptx_position:
            pptx_position = self.default_position
        kwargs.update(pptx_position.dict())
        with io.BytesIO() as output:
            fig.savefig(output, format="png")
            pic = slide.shapes.add_picture(output, **kwargs)  # 0, 0)#, left, top)
        return pic

    def add_text_box(self, slide, text: str, position: PPTXPosition = None, font: PPTXFontStyle = None) -> Shape:
        """
        Adds a text box with given text using given position and font.
        Uses self.default_position if no position is given.
        """
        width = height = Inches(1)  # no auto-resizing of shape -> has to be done inside PowerPoint
        if position is None:
            position = self.default_position
        result = slide.shapes.add_textbox(**position.dict(), width=width, height=height)
        result.text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
        result.text_frame.text = text  # first paragraph
        if font:
            font.write_shape(result)
        return result

    def move_slide(self, slide: Slide, new_index: int):
        """Moves the given slide to position new_index."""
        _sldIdLst = self.prs.slides._sldIdLst

        old_index = None
        for index, entry in enumerate((_sldIdLst.sldId_lst)):
            if entry.id == slide.slide_id:
                old_index = index

        if old_index is not None:
            to_move = _sldIdLst[old_index]
            list(_sldIdLst).pop(old_index)
            _sldIdLst.insert(new_index, to_move)

    @staticmethod
    def remove_unpopulated_shapes(slide: Slide):
        """
        Removes empty placeholders (e.g. due to layout) from slide.
        Further testing needed.
        """
        for index in reversed(range(len(slide.shapes))):
            shape = slide.shapes[index]
            # if shape.is_placeholder and shape.text_frame.text == "":
            if shape.has_text_frame and shape.text_frame.text == "":
                shape.element.getparent().remove(shape.element)
                print(f"removed index {index}")

    @staticmethod
    def create_hyperlink(run: _Run, shape: Shape, to_slide: Slide):  # text hyperlink not implemented in pptx-python
        """Makes the given run a hyperlink to to_slide."""
        shape.click_action.target_slide = to_slide
        run.hyperlink.address = shape.click_action.hyperlink.address
        run.hyperlink._hlinkClick.action = shape.click_action.hyperlink._hlink.action
        run.hyperlink._hlinkClick.rId = shape.click_action.hyperlink._hlink.rId
        shape.click_action.target_slide = None

    def add_content_slide(self, slide_index=1):
        """Adds a content slide with hyperlinks to all other slides and puts it to position slide_index."""
        content_entries = []

        for slide in self.prs.slides:
            content_entries.append((slide.shapes.title.text, slide))

        result = self.add_slide("Content")
        content_text_box = self.add_text_box(result, "", PPTXPosition(0.1, 0.2))
        for text, slide in content_entries:
            paragraph = content_text_box.text_frame.add_paragraph()
            run = paragraph.add_run()
            run.text = text
            self.create_hyperlink(run, content_text_box, slide)

        self.move_slide(result, slide_index)

        return result
