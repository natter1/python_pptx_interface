"""
This module is a collection of helpful misc. functions.
@author: Nathanael Jöhrmann
"""
# from pptx.presentation import Presentation
import _ctypes
import os
from typing import Generator

from comtypes.client import Constants, CreateObject

import pptx
import tempfile


class TempFileGenerator:
    generator = None
    @classmethod
    def get_new_generator(cls, prs):
        TempFileGenerator.generator = cls.temp_generator(prs)
        return cls.generator

    @staticmethod
    def temp_generator(prs: pptx.presentation.Presentation) -> Generator[str, None, None]:
        temp_pptx = None
        try:
            with tempfile.NamedTemporaryFile(suffix='.pptx', delete=False) as temp_pptx:
                temp_pptx.close()  # needed on windows systems to access file
                prs.save(temp_pptx.name)
                yield temp_pptx.name
        finally:
            if temp_pptx is not None:
                try:
                    os.remove(temp_pptx.name)
                except PermissionError as e:  # file still in use somewhere
                    pass


def get_temporary_pptx_filename(prs: pptx.presentation.Presentation) -> str:
    """
    Generates a temporary pptx file. This is useful under windows, where tempfile.NamedTemporaryFile is broken.
    Yields the filename of the temporary file.
    """
    my_generator = TempFileGenerator.get_new_generator(prs)
    for filename in my_generator:
        return filename


# ----------------------------------------------------------------------------------------------------------------------
# The following functions need an installed PowerPoint and will only work on windows systems.
# ----------------------------------------------------------------------------------------------------------------------
def save_pptx_as_png(png_filename: str, pptx_filename, overwrite_folder: bool = False):
    if os.path.isdir(png_filename) and not overwrite_folder:
        print(f"Folder {png_filename} already exists. "
              f"Set overwrite_folder=True, if you want to overwrite folder content.")
        return False

    powerpoint = CreateObject("Powerpoint.Application")
    pp_constants = Constants(powerpoint)

    pres = powerpoint.Presentations.Open(pptx_filename)
    pres.SaveAs(png_filename, pp_constants.ppSaveAsPNG)
    pres.close()
    if powerpoint.Presentations.Count == 0:  # only close, when no other Presentations are open!
        powerpoint.quit()


def save_pptx_as_pdf(pdf_filename: str, pptx_filename, overwrite: bool = False) -> bool:
    """
    :param pdf_filename: filename (including path) of new pdf file
    :param pptx_filename: filename (including path) of pptx file
    :return:
    """
    if os.path.isfile(pdf_filename) and not overwrite:
        print(f"File {pdf_filename} already exists. Set overwrite=True, if you want to overwrite file.")
        return False

    powerpoint = CreateObject("Powerpoint.Application")
    pp_constants = Constants(powerpoint)
    pres = powerpoint.Presentations.Open(pptx_filename)
    pres.SaveAs(pdf_filename, pp_constants.ppSaveAsPDF)
    pres.close()
    if powerpoint.Presentations.Count == 0:  # only close, when no other Presentations are open!
        powerpoint.quit()

    return True


def save_as_pdf(prs: pptx.presentation.Presentation, filename: str, overwrite: bool = False) -> bool:
    """
    Save presentation as PDF.
    Requires to save a temporary *.pptx first.
    Needs module comtypes (windows only).
    Needs installed PowerPoint.
    Note: you have to give full path for filename, or PowerPoint might cause random exceptions.
    """
    result = False
    try:
        result = save_pptx_as_pdf(filename, get_temporary_pptx_filename(prs), overwrite)
    except _ctypes.COMError as e:
        print(e)
        print("Couldn't save PDF file due to communication error with PowerPoint.")
        result = False

    return result


def save_as_png(prs: pptx.presentation.Presentation, filename: str, overwrite: bool = False) -> bool:
    """
    Save presentation as PDF.
    Requires to save a temporary *.pptx first.
    Needs module comtypes (windows only).
    Needs installed PowerPoint.
    Note: you have to give full path for filename, or PowerPoint might cause random exceptions.
    """
    result = False
    try:
        result = save_pptx_as_png(filename, get_temporary_pptx_filename(prs), overwrite)
    except _ctypes.COMError as e:
        print(e)
        print("Couldn't save PNG file due to communication error with PowerPoint.")
        result = False

    return result
