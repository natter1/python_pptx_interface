import pytest
from pptx.util import Inches

from pptx_tools.creator import PPTXCreator
from pptx_tools.templates import TemplateExample


@pytest.fixture(scope='class')
def pptx_creator():
    creator = PPTXCreator()
    yield creator

# todo: put in fixture to call test several times
fraction_test_cases = [-4, 0, 0.4, 1, 2]

class TestPPTXCreator:
    def test__fraction_width_to_inch(self, pptx_creator):
        for fraction in fraction_test_cases:
            inches_calc = pptx_creator._fraction_width_to_inch(fraction)
            inches_from_prs = pptx_creator.prs.slide_width * fraction
            assert type(inches_calc) is Inches  # type is Inches even if data is int!
            assert inches_calc == inches_from_prs

    def test__fraction_height_to_inch(self, pptx_creator):
        for fraction in fraction_test_cases:
            inches_calc = pptx_creator._fraction_height_to_inch(fraction)
            inches_from_prs = pptx_creator.prs.slide_height * fraction
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

    def test__write_position_in_kwargs(self):
        assert False

    def test_add_matplotlib_figure(self):
        assert False

    def test_add_text_box(self):
        assert False

    def test__get_rows_cols(self):
        assert False

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
