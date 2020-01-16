"""
This module provides an easier Interface to create *.pptx presentations using the module python-pptx.
@author: Nathanael Jöhrmann
"""
import io

from pptx import Presentation
from pptx.shapes.autoshape import Shape
from pptx.util import Inches
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.slide import Slide

from pptx_tools.font_style import PPTXFontStyle


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
        - default_position elements as fraction of slide height/width
        - add matplotlib figures
        - use pptx templates (in combination with templates.py)
    """

    def __init__(self, template=None):
        """
        :param template:
        :param title:
        """
        self.slides = []
        self.template = None
        self.prs = None
        self.title_layout = None
        self.default_layout = None
        self.create_presentation(template)
        self.default_position = PPTXPosition(presentation=self.prs)


    def fraction_width_to_inch(self, fraction: float) -> Inches:
        """
        Returns a width in inches calculated as a fraction of total slide-width.
        :param fraction: float
        :return: Calculated Width in inch
        """
        result = Inches(self.prs.slide_width.inches * fraction)
        return result

    def fraction_height_to_inch(self, fraction: float) -> Inches:
        """
        Returns a height in inches calculated as a fraction of total slide-height.
        :param fraction: float
        :return: Calculated Width in inch
        """
        return Inches(self.prs.slide_height.inches * fraction)

    def save(self, filename: str = "delme.pptx") -> None:
        """
        Saves the presentation under the given filename.
        :param filename: string
        :return: None
        """
        self.prs.save(filename)

    def create_presentation(self, template=None):
        """
        Create a new presentation (using optional template)
        """
        if template:
            self.create_presentation_from_template(template)
        else:
            self.prs = Presentation()
            self.title_layout = self.prs.slide_masters[0].slide_layouts[0]
            self.default_layout = self.prs.slide_masters[0].slide_layouts[0]

    def create_presentation_from_template(self, template) -> None:
        self.template = template
        self.prs = template.prs
        self.title_layout = template.title_layout
        self.default_layout = template.default_layout

    def create_title_slide(self, title, layout=None) -> Slide:
        if not layout:
            layout = self.title_layout
        return self.add_slide(title, layout)

    def add_slide(self, title, layout=None) -> Slide:
        if not layout:
            layout = self.default_layout
        slide = self.prs.slides.add_slide(layout)
        title_shape = slide.shapes.title
        title_shape.text = title
        self.remove_unpopulated_shapes(slide)
        return slide

    def write_position_in_kwargs(self, left_rel=0.0, top_rel=0.0, kwargs={}):
        """
        This method modifies(!) the argument kwargs by adding or changing the entries "left" and "top".
        :param left_rel:
        :param top_rel:
        :param kwargs:
        :return: None
        """
        left = self.fraction_width_to_inch(left_rel)
        top = self.fraction_height_to_inch(top_rel)

        if not "left" in kwargs:
            kwargs["left"] = left
        else:
            kwargs["left"] = kwargs["left"] + left
        if not "top" in kwargs:
            kwargs["top"] = top
        else:
            kwargs["top"] = kwargs["top"] + top

    def add_matplotlib_figure(self, fig, slide, pptx_position: PPTXPosition = None, zoom: float = 1.0, **kwargs):
        """
        Add a motplotlib figure fig to slide with index slide_index. With top_rel and left_rel
        it is possible to default_position the figure in Units of slide height/width (float in range [0, 1].
        :param pptx_position: PPTXPosition
        :param fig: a matplolib figure
        :param slide: slide in presentation on which to insert fig
        :param zoom: sets image scaling in PowerPoint; only used if width not in kwargs (default = 1.0)
        :param kwargs:
        :return: pptx.shapes.picture.Picture
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

    def add_text_box(self, slide, text: str, position: PPTXPosition, font: PPTXFontStyle = None) -> Shape:
        """ Adds a tex box with given text using given position and font."""
        width = height = Inches(1)  # no auto-resizing of shape -> has to be done inside PowerPoint
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
    def remove_unpopulated_shapes(slide):
        """
        Removes empty placeholders (e.g. due to layout) from slide.
        Further testing needed.
        :param slide: pptx.slide.Slide
        :return:
        """
        for index in reversed(range(len(slide.shapes))):
            shape = slide.shapes[index]
            # if shape.is_placeholder and shape.text_frame.text == "":
            if shape.has_text_frame and shape.text_frame.text == "":
                shape.element.getparent().remove(shape.element)
                print(f"removed index {index}")

    def add_content_slide(self, slide_index=1):
        def create_hyperlink(run, shape, to_slide):  # text hyperlink not implemented in pptx-python; hack via shape
            shape.click_action.target_slide = to_slide
            run.hyperlink.address = shape.click_action.hyperlink.address
            run.hyperlink._hlinkClick.action = shape.click_action.hyperlink._hlink.action
            run.hyperlink._hlinkClick.rId = shape.click_action.hyperlink._hlink.rId
            shape.click_action.target_slide = None

        content_entries = []

        for slide in self.prs.slides:
            content_entries.append((slide.shapes.title.text, slide))

        result = self.add_slide("Content")
        content_text_box = self.add_text_box(result, "", PPTXPosition(0.1, 0.2))
        for text, slide in content_entries:
            paragraph = content_text_box.text_frame.add_paragraph()
            run = paragraph.add_run()
            run.text = text
            create_hyperlink(run, content_text_box, slide)

        self.move_slide(result, slide_index)

        return result
