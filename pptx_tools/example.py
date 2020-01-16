from pptx import Presentation

from pptx_tools.creator import PPTXCreator
from pptx_tools.style_sheets import font_default, font_title, font_slide_title, font_sub_title
from pptx_tools.templates import TemplateExample

def test():
    pp = PPTXCreator()
    pp.prs = Presentation("example.pptx")
    content_slide = pp.prs.slides[1]
    for shape in content_slide.shapes:
        print(shape)

def run():
    #test()

    pp = PPTXCreator(TemplateExample())

    title_slide = pp.create_title_slide("Example presentation")
    font = font_title()
    font.write_shape(title_slide.shapes.title)


    pp.add_slide("page2")
    pp.add_slide("page3")
    slide = pp.add_slide("page4")

    content_slide = pp.add_content_slide()
    # pp.move_slide(content_slide, 1)
    # _sldIdLst = pp.prs.slides._sldIdLst
    #
    # move_entry = _sldIdLst.sldId_lst[3]
    # move_entry = _sldIdLst[3]
    # try:
    #     # _sldIdLst.sldId_lst.remove(move_entry)
    #     # _sldIdLst.sldId_lst = []
    #     list(_sldIdLst).pop(3)
    #     _sldIdLst.insert(0, move_entry)
    # except ValueError:
    #     raise Exception

    pp.save("example.pptx")


if __name__ == '__main__':
    run()