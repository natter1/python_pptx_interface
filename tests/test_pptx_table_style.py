"""
@author: Nathanael Jöhrmann
"""
import pytest

from pptx_tools.creator import PPTXCreator
from pptx_tools.templates import TemplateExample


@pytest.fixture(scope='class')
def pptx_creator():
    creator = PPTXCreator(TemplateExample())
    yield creator

class TestPPTXTableStyle:
    def test_read_table(self):
        assert False

    def test_set(self):
        assert False

    def test__write_all_cells(self):
        assert False

    def test__update_col_ratios(self):
        assert False

    def test__write_col_sizes(self):
        assert False

    def test_write_shape(self):
        assert False

    def test_write_table(self):
        assert False

    def test_set_width_as_fraction(self):
        assert False
