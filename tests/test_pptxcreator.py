import matplotlib.pyplot as plt
import pytest
from pptx.util import Inches

from pptx_tools.creator import PPTXCreator
from pptx_tools.position import PPTXPosition
from pptx_tools.templates import TemplateExample


@pytest.fixture(scope='class')
def pptx_creator():
    creator = PPTXCreator()
    yield creator


@pytest.fixture(scope="function", params=[-4, 0, 0.4, 1, 2])
def fraction_test_case(request):
    yield request.param


@pytest.fixture(scope='class')
def matplotlib_figure():
    yield plt.figure(figsize=(2, 1), dpi=100, facecolor='w', edgecolor='w', frameon=True)


class TestPPTXCreator:
    def test__fraction_width_to_inch(self, pptx_creator, fraction_test_case):
        inches_calc = pptx_creator._fraction_width_to_inch(fraction_test_case)
        inches_from_prs = pptx_creator.prs.slide_width * fraction_test_case
        assert type(inches_calc) is Inches  # type is Inches even if data is int!
        assert inches_calc == inches_from_prs

    def test__fraction_height_to_inch(self, pptx_creator, fraction_test_case):
        inches_calc = pptx_creator._fraction_height_to_inch(fraction_test_case)
        inches_from_prs = pptx_creator.prs.slide_height * fraction_test_case
        assert type(inches_calc) is Inches  # type is Inches even if data is int!
        assert inches_calc == inches_from_prs

    def test__create_presentation(self, pptx_creator):
        old_prs = pptx_creator.prs
        pptx_creator._create_presentation()
        new_prs = pptx_creator.prs
        assert old_prs is not new_prs
        assert new_prs is not None
        pptx_creator._create_presentation(TemplateExample())
        new_prs_with_template = pptx_creator.prs
        assert new_prs is not new_prs_with_template
        assert new_prs_with_template is not None

    def test__create_presentation_from_template(self, pptx_creator):
        old_prs = pptx_creator.prs
        pptx_creator._create_presentation_from_template(TemplateExample())
        new_prs_with_template = pptx_creator.prs
        assert old_prs is not new_prs_with_template
        assert new_prs_with_template is not None

    def test_add_title_slide(self, pptx_creator):
        n_slides_before = len(pptx_creator.prs.slides)
        pptx_creator.add_title_slide("Title slide")
        n_slides_after = len(pptx_creator.prs.slides)
        assert n_slides_before == n_slides_after - 1

    def test_add_slide(self, pptx_creator):
        n_slides_before = len(pptx_creator.prs.slides)
        pptx_creator.add_slide("slide")
        n_slides_after = len(pptx_creator.prs.slides)
        assert n_slides_before == n_slides_after - 1

    def test_add_matplotlib_figure(self, pptx_creator, matplotlib_figure):
        slide = pptx_creator.add_slide("slide")
        fig_width = matplotlib_figure.bbox_inches.width
        fig_height = matplotlib_figure.bbox_inches.height
        zoom = 0.8
        position = PPTXPosition(0.6, 0.4, 1, -1)
        shape = pptx_creator.add_matplotlib_figure(matplotlib_figure, slide, position, zoom=zoom)

        assert pptx_creator.prs.slide_width * position.left_rel + Inches(position.left) == shape.left
        assert pptx_creator.prs.slide_height * position.top_rel + Inches(position.top) == shape.top
        assert fig_width * zoom == shape.width.inches
        assert fig_height * zoom == shape.height.inches

    def test_add_text_box(self, pptx_creator):
        slide = pptx_creator.add_slide("slide")
        position = PPTXPosition(0.6, 0.4, 1, -1)
        text = "Test text"

        shape = pptx_creator.add_text_box(slide, text, position)

        assert pptx_creator.prs.slide_width * position.left_rel + Inches(position.left) == shape.left
        assert pptx_creator.prs.slide_height * position.top_rel + Inches(position.top) == shape.top
        assert text == shape.text

    def test__get_rows_cols(self, pptx_creator):
        table_data = [[0, 1, 2], [1], [2], [3], [4]]  # 5 rows; 3 cols
        result = pptx_creator._get_rows_cols(table_data)
        assert result == (5, 3)

    def test_add_table(self):
        assert False

    def test_move_slide(self):
        assert False

    def test_remove_unpopulated_shapes(self):
        assert False

    def test_create_hyperlink(self):
        assert False

    def test_add_content_slide(self):
        assert False

    def test_save(self):
        assert False

    def test_save_as_pdf(self):
        assert False

    def test_save_as_png(self):
        assert False
