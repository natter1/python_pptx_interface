"""
This file is a collection of PowerPoint enumerations still missing in python-pptx.
@author: Nathanael Jöhrmann
"""
from enum import Enum


class TEXT_STRIKE_VALUES(Enum):
    """
    https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textstrikevalues
    """
    DoubleStrike = "dblStrike"
    NoStrike = "noStrike"
    SingleStrike = "sngStrike"


class TEXT_CAPS_VALUES(Enum):
    """
    https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textcapsvalues
    """
    All = "all"
    None_ = "none"
    Small = "small"
