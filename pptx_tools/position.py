"""
@author: Nathanael Jöhrmann
"""
from typing import Optional

from pptx.util import Inches


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

    def set(self, left_rel=0.0, top_rel=0.0, left=0, top=0):
        self.left_rel = left_rel
        self.top_rel = top_rel
        self.left = left
        self.top = top

    def _dict_for_position(self, left_rel=0.0, top_rel=0.0, left=0, top=0):
        """
        Returns kwargs dict for given default_position. Does not change attributes of self
        :param left_rel: float [slide_width]
        :param top_rel: float [slide_height]
        :param left: float [inch]
        :param top: float [inch]
        :return: dictionary
        """
        left = self._fraction_width_to_inch(left_rel) + Inches(left)
        top = self._fraction_height_to_inch(top_rel) + Inches(top)
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

    def _fraction_width_to_inch(self, fraction):
        """
        Returns a width in inches calculated as a fraction of total slide-width.
        :param fraction: float
        :return: Calculated Width in inch
        """
        result = Inches(self.prs.slide_width.inches * fraction)
        return result

    def _fraction_height_to_inch(self, fraction):
        """
        Returns a height in inches calculated as a fraction of total slide-height.
        :param fraction: float
        :return: Calculated Width in inch
        """
        return Inches(self.prs.slide_height.inches * fraction)