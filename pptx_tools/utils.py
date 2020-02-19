"""
This module is a collection of helpful misc. functions.
@author: Nathanael Jöhrmann
"""
import _ctypes
import os
from typing import Generator, Union


try:
    from comtypes.client import Constants, CreateObject
    has_comptypes=True
except:
    has_comptypes=False

import pptx
from pptx.table import Table, _Cell
import tempfile


class TemporaryPPTXFile:
    __slots__ = ('_file', 'dir', 'filepath', 'raise_on_delete_error')

    def __init__(self, mode="w+b", suffix=".pptx", dir=None, raise_on_delete_error = True):
        if not dir:
            dir = tempfile.gettempdir()
        self.dir = dir
        self.filepath = os.path.join(dir, os.urandom(32).hex() + suffix)
        self._file = open(self.filepath, mode)
        self.raise_on_delete_error = raise_on_delete_error

    def __enter__(self):
        return self._file.__enter__()

    def __exit__(self, exc_type, exc_value, traceback):
        ret = self._file.__exit__(exc_type, exc_value, traceback)
        try:
            os.remove(self._file.name)
        except PermissionError as e:
            if self.raise_on_delete_error:
                raise PermissionError(e)
            else:
                print(e)
        return ret


class _USE_DEFAULT:  # using a class allows typing
    def __str__(self):
        return "This is a default value, used to express that a value should become default, which is indicated with " \
               "None in python-pptx. But in python-pptx-interface styles None generally means 'do not change'. " \
               "An example would be 'PPTXFontStyle.size = None'. This would ensure, that the font size will not be " \
               "changed when calling PPTXFontStyle.write_font(). But to remove a customized font size, e.g. in a run, " \
               "the value has to be set to None in python-pptx. Thats done with 'PPTXFontStyle.size = use_default'."


def use_default():
    return _USE_DEFAULT


def iter_table_cells(table:  Table) -> Generator[_Cell, None, None]:
    for row in table.rows:
        for cell in row.cells:
            yield cell


def change_paragraph_text_to(paragraph, text):
    """
    Change text of paragraph to text, but keep format of first run.
    :param paragraph:
    :param text:
    :return:
    """
    from pptx_tools.font_style import PPTXFontStyle  # local import to prevent circle import error
    font = PPTXFontStyle()
    font.read_font(paragraph.runs[0].font)
    paragraph.text = text
    font.write_paragraph(paragraph)


def copy_font(_from: 'Font', _to: 'Font') -> None:
    """Copies settings from one pptx.text.text.Font to another."""
    from pptx_tools.font_style import PPTXFontStyle  # local import to prevent circle import error
    font_style = PPTXFontStyle()
    font_style.read_font(_from)
    font_style.write_font(_to)

# ----------------------------------------------------------------------------------------------------------------------
# The following functions need an installed PowerPoint and will only work on windows systems.
# ----------------------------------------------------------------------------------------------------------------------
def save_pptx_as_png(save_folder: Union[str, "LocalPath"], pptx_filename: str, overwrite_folder: bool = False) -> bool:
    if not has_comptypes:
        print("Comptype module needed to save PNGs.")
        return False

    if os.path.isdir(save_folder) and not overwrite_folder:
        print(f"Folder {save_folder} already exists. "
              f"Set overwrite_folder=True, if you want to overwrite folder content.")
        return False

    powerpoint = CreateObject("Powerpoint.Application")
    pp_constants = Constants(powerpoint)

    pres = powerpoint.Presentations.Open(pptx_filename)
    pres.SaveAs(str(save_folder), pp_constants.ppSaveAsPNG)
    pres.close()
    if powerpoint.Presentations.Count == 0:  # only close, when no other Presentations are open!
        powerpoint.quit()
    return True

def save_pptx_as_pdf(pdf_filename: Union[str, "LocalPath"], pptx_filename, overwrite: bool = False) -> bool:
    """
    :param pdf_filename: save_folder (including path) of new pdf file
    :param pptx_filename: save_folder (including path) of pptx file
    :return:
    """
    if not has_comptypes:
        print("Comptype module needed to save PDFs.")
        return False

    if os.path.isfile(pdf_filename) and not overwrite:
        print(f"File {pdf_filename} already exists. Set overwrite=True, if you want to overwrite file.")
        return False

    powerpoint = CreateObject("Powerpoint.Application")
    pp_constants = Constants(powerpoint)
    pres = powerpoint.Presentations.Open(pptx_filename)
    pres.SaveAs(str(pdf_filename), pp_constants.ppSaveAsPDF)
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
    Note: you have to give full path for save_folder, or PowerPoint might cause random exceptions.
    """
    result = False
    with TemporaryPPTXFile() as f:
        prs.save(f.name)
        try:
            result = save_pptx_as_pdf(filename, f.name, overwrite)
        except _ctypes.COMError as e:
            print(e)
            print("Couldn't save PDF file due to communication error with PowerPoint.")
            result = False
    return result


def save_as_png(prs: pptx.presentation.Presentation, save_folder: str, overwrite: bool = False) -> bool:
    """
    Save presentation as PDF.
    Requires to save a temporary *.pptx first.
    Needs module comtypes (windows only).
    Needs installed PowerPoint.
    Note: you have to give full path for save_folder, or PowerPoint might cause random exceptions.
    """
    result = False
    with TemporaryPPTXFile() as f:
        prs.save(f.name)
        try:
            result = save_pptx_as_png(save_folder, f.name, overwrite)
        except _ctypes.COMError as e:
            print(e)
            print("Couldn't save PNG file due to communication error with PowerPoint.")
            result = False
    return result
