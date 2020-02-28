"""
This file contains tests for PPTXPosition-methods.
@author: Nathanael Jöhrmann
"""
import pytest
from pptx.util import Pt, Emu, Inches

from pptx_tools.creator import PPTXCreator
from pptx_tools.position import PPTXPosition
from pptx_tools.templates import TemplateExample


@pytest.fixture(scope='session')
def pptx_creator():
    creator = PPTXCreator(TemplateExample())
    yield creator


@pytest.fixture(scope='function')
def pptx_position(pptx_creator):
    position = PPTXPosition(0, 0.5, 1, 2)
    yield position


class TestPPTXPosition:
    def test__eq__(self, pptx_position):
        assert pptx_position == PPTXPosition(0, 0.5, 1, 2)
        assert PPTXPosition(0.2, 0, 2, 1) == PPTXPosition(0, 0, Emu(4267200).inches,  Emu(914400).inches)

    def test_set(self, pptx_position):
        pptx_position.set()  # all arguments optional - without arguments should not change anything!
        assert pptx_position == PPTXPosition(0, 0.5, 1, 2)
        pptx_position.set(left_rel=0.2, top_rel=0, left=2, top=1)
        assert pptx_position == PPTXPosition(0.2, 0, 2, 1)

    def test__dict_for_position(self, pptx_position):
        assert PPTXPosition._dict_for_position(1, 2, 3, 4) == {'left': 14935200, 'top': 17373600}

    def test_dict(self, pptx_position):
        assert pptx_position.dict() == {'left': 914400, 'top': 5257800}

    def test_tuple(self, pptx_position):
        assert pptx_position.tuple() == (914400, 5257800)

    def test__fraction_width_to_inch(self, pptx_position):
        assert PPTXPosition()._fraction_width_to_inch(0) == Inches(0)
        assert PPTXPosition()._fraction_width_to_inch(1) == Inches(13.0 + 1/3)

    def test__fraction_height_to_inch(self, pptx_position):
        assert PPTXPosition()._fraction_height_to_inch(0) == Inches(0)
        assert PPTXPosition()._fraction_height_to_inch(1) == Inches(7.5)
