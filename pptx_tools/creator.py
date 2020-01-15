"""
This module provides an easier Interface to create *.pptx presentations using the module python-pptx.
@author: Nathanael Jöhrmann
"""
import io

from pptx import Presentation
from pptx.util import Inches
from pptx.enum.text import MSO_AUTO_SIZE


def main():
    pass
    # from pptx_tools.templates import TemplateExample
    # from pptx_tools.pi88_importer import PI88Measurement
    # from pptx_tools.pi88_plotter import PI88Plotter
    #
    # # pptx = PPTXCreator()  # create pptx without using a template file
    # pptx = PPTXCreator(TemplateExample())
    # title_slide = pptx.create_title_slide(title="PPTXCreator Demo")
    # slide2 = pptx.add_slide(title="Normal slide")
    #
    # plotter = PI88Plotter(PI88Measurement("..\\resources\\AuSn_Creep\\1000uN 01 LC.tdm"))
    # fig = plotter.get_load_displacement_plot()
    # fig_width = fig.get_figwidth()
    # zoom = 1
    # picture = pptx.add_matplotlib_figure(fig, title_slide, PPTXPosition(0.2, 0.25),
    #                                      width=Inches(fig_width * zoom))
    #
    #
    # zoom = 0.2
    # pptx.add_matplotlib_figure(fig, title_slide, PPTXPosition(0.7, 0.25), width=Inches(fig_width * zoom))
    #
    # pptx.add_text_box(slide2, "This is the first paragraph\n... and here comes the second", PPTXPosition(0.7, 0.7))
    # pptx.save("delme_pptx_creator_demo.pptx")


class PPTXPosition:
    # reposition picture:
    # picture.left = Inches(1)
    # picture.top = Inches(3)
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

    def create_presentation_from_template(self, template):
        self.template = template
        self.prs = template.prs
        self.title_layout = template.title_layout
        self.default_layout = template.default_layout

    def create_title_slide(self, title, layout=None):
        if not layout:
            layout = self.title_layout
        return self.add_slide(title, layout)

    def add_slide(self, title, layout=None):
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
            kwargs["width"] = fig.get_figwidth() * zoom
        if not pptx_position:
            pptx_position = self.default_position
        kwargs.update(pptx_position.dict())
        with io.BytesIO() as output:
            fig.savefig(output, format="png")
            pic = slide.shapes.add_picture(output, **kwargs)  # 0, 0)#, left, top)
        return pic

    def add_text_box(self, slide, text: str, position: PPTXPosition):  # -> added text box shape
        # todo: implement font
        width = height = Inches(1)  # no auto-resizing of shape -> has to be done inside PowerPoint
        result = slide.shapes.add_textbox(**position.dict(), width=width, height=height)
        result.text_frame.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
        result.text_frame.text = text  # first paragraph
        # todo: remove - only for testing:
        # result.text_frame.add_paragraph().text = "aditional paragraph"
        return result

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


if __name__ == '__main__':
    main()
