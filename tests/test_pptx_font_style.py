"""
This file contains tests for PPTXFontStyle-methods. It can be useful to open the created pptx-file in the local temp
folder, to also check functionality inside PowerPoint.
@author: Nathanael Jöhrmann
"""
import os

import pytest

from pptx_tools.creator import PPTXCreator
from pptx_tools.templates import TemplateExample


@pytest.fixture(scope='session')
def pptx_creator():
    creator = PPTXCreator(TemplateExample())
    yield creator


class TestPPTXFontStyle:
    def test_color_rgb(self):
        assert False

    def test_color_rgb(self):
        assert False

    def test_read_font(self):
        assert False

    def test_write_font(self):
        assert False

    def test__write_caps(self):
        assert False

    def test__write_strikethrough(self):
        assert False

    def test__get_write_value(self):
        assert False

    def test_write_shape(self):
        assert False

    def test_write_text_frame(self):
        assert False

    def test_write_paragraph(self):
        assert False

    def test_write_run(self):
        assert False

    def test_set(self):
        assert False

def test_save_test_results_as_temp_pptx_file(pptx_creator, tmpdir):
    file = tmpdir.join("test_font_style.pptx")
    pptx_creator.save(file)
    assert True
