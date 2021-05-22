"""
@author: Nathanael Jöhrmann
"""
from typing import Optional

from pptx.util import Inches

from pptx_tools.utils import _DO_NOT_CHANGE


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

    def __init__(self, left_rel=0.0, top_rel=0.0, left=0, top=0, presentation: Optional["PPTXCreator"] = None):
        """
        :param presentation: pptx.prs (needed for slide width and height)
        :param left_rel: distance from slide left (relative to slide width)
        :param top_rel: distance from slide top (relative to slide height)
        :param left: left position [inches] starting from rel_left
        :param top: top position [inches] starting from rel_top
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

    def set(self, left_rel=_DO_NOT_CHANGE, top_rel=_DO_NOT_CHANGE, left=_DO_NOT_CHANGE, top=_DO_NOT_CHANGE):
        """Convenience method to set several PPTXPosition attributes together."""
        if left_rel is not _DO_NOT_CHANGE:
            self.left_rel = left_rel
        if top_rel is not _DO_NOT_CHANGE:
            self.top_rel = top_rel
        if left is not _DO_NOT_CHANGE:
            self.left = left
        if top is not _DO_NOT_CHANGE:
            self.top = top

    @classmethod
    def _dict_for_position(cls, left_rel=0.0, top_rel=0.0, left=0, top=0):
        """
        Returns kwargs dict for given default_position. Does not change attributes of self
        :param left_rel: float [slide_width]
        :param top_rel: float [slide_height]
        :param left: float [inch]
        :param top: float [inch]
        :return: dictionary
        """
        left = cls._fraction_width(left_rel) + Inches(left)
        top = cls._fraction_height(top_rel) + Inches(top)
        return {"left": left, "top": top}

    def dict(self):
        """
        This method returns a kwargs dict containing "left" and "top".
        :return: dictionary
        """
        return self._dict_for_position(self.left_rel, self.top_rel, self.left, self.top)

    def tuple(self):
        """
        This method returns an args tuple containing "left" and "top".
        :return: tuple
        """
        left = self.dict()["left"]
        top = self.dict()["top"]
        return left, top

    @classmethod
    def _fraction_width(cls, fraction):
        """
        Returns a width in pptx units (integer) calculated as a fraction of total slide-width.
        :param fraction: float
        :return: Calculated Width in inch
        """
        return Inches(cls.prs.slide_width.inches) * fraction

    @classmethod
    def _fraction_height(cls, fraction):
        """
        Returns a height in pptx units (integer) calculated as a fraction of total slide-height.
        :param fraction: float
        :return: Calculated Width in inch
        """
        return Inches(cls.prs.slide_height.inches) * fraction

    @classmethod
    def _fraction_width_to_inch(cls, fraction):
        """
        Returns a width in inches calculated as a fraction of total slide-width.
        :param fraction: float
        :return: Calculated Width in inch
        """
        return cls.prs.slide_width.inches * fraction

    @classmethod
    def _fraction_height_to_inch(cls, fraction):
        """
        Returns a height in inches calculated as a fraction of total slide-height.
        :param fraction: float
        :return: Calculated Width in inch
        """
        return cls.prs.slide_height.inches * fraction

    def __eq__(self, other):
        """Overrides the default implementation"""
        if isinstance(other, PPTXPosition):
            print(f"self.dict(): {self.dict()} - other.dict(): {other.dict()}")
            return self.dict() == other.dict()
        return NotImplemented
