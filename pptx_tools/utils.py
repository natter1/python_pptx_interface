"""
This module is a collection of helpful misc. functions.
@author: Nathanael Jöhrmann
"""
# from pptx.presentation import Presentation
import _ctypes
import os
from typing import Generator

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


def iter_table_cells(table:  Table) -> Generator[_Cell, None, None]:
    for row in table.rows:
        for cell in row.cells:
            yield cell

# ----------------------------------------------------------------------------------------------------------------------
# The following functions need an installed PowerPoint and will only work on windows systems.
# ----------------------------------------------------------------------------------------------------------------------
def save_pptx_as_png(png_foldername: str, pptx_filename: str, overwrite_folder: bool = False):

    if not has_comptypes:
        print("Comptype module needed to save PDFs.")
        return

    if os.path.isdir(png_foldername) and not overwrite_folder:
        print(f"Folder {png_foldername} already exists. "
              f"Set overwrite_folder=True, if you want to overwrite folder content.")
        return

    powerpoint = CreateObject("Powerpoint.Application")
    pp_constants = Constants(powerpoint)

    pres = powerpoint.Presentations.Open(pptx_filename)
    pres.SaveAs(png_foldername, pp_constants.ppSaveAsPNG)
    pres.close()
    if powerpoint.Presentations.Count == 0:  # only close, when no other Presentations are open!
        powerpoint.quit()


def save_pptx_as_pdf(pdf_filename: str, pptx_filename, overwrite: bool = False) -> bool:
    """
    :param pdf_filename: filename (including path) of new pdf file
    :param pptx_filename: filename (including path) of pptx file
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
    with TemporaryPPTXFile() as f:
        prs.save(f.name)
        try:
            result = save_pptx_as_pdf(filename, f.name, overwrite)
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
    with TemporaryPPTXFile() as f:
        prs.save(f.name)
        try:
            result = save_pptx_as_png(filename, f.name, overwrite)
        except _ctypes.COMError as e:
            print(e)
            print("Couldn't save PNG file due to communication error with PowerPoint.")
            result = False
    return result
