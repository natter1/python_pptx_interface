from pptx_tools.creator import PPTXCreator, PPTXPosition
from pptx_tools.style_sheets import font_title
from pptx_tools.templates import TemplateExample

try:
    import matplotlib.pyplot as plt
    matplotlib_installed = True
except ImportError as e:
    matplotlib_installed = False


def create_demo_figure():
    figure: plt.Figure = plt.figure(figsize=(3.4, 1.8), dpi=100, facecolor='w', edgecolor='w', frameon=True)
    # figure.patch  # The Rectangle instance representing the figure background patch.
    # figure.patch.set_visible(False)
    figure.patch.set_alpha(0.5)
    sub_title = figure.suptitle('matplotlib figure', fontsize=14, fontweight='bold', color='red')
    sub_title.set_color('green')
    sub_title.set_rotation(5)
    sub_title.set_size(18)

    textstr = '\n'.join((
        fr'$\mu={5}^{5}$',
        r'$\mathrm{median}=3_3$',
        r'$median=3_3$',
        fr'$\sigma=$'))
    figure.text(0.3, 0.05, textstr)
    return figure


def run():
    pp = PPTXCreator(TemplateExample())

    title_slide = pp.add_title_slide("Example presentation")
    font = font_title()
    font.write_shape(title_slide.shapes.title)  # you can change font attributes of paragraphs in shape via PPTXFontTool

    slide2 = pp.add_slide("page2")
    pp.add_slide("page3")
    pp.add_slide("page4")
    content_slide = pp.add_content_slide()

    if matplotlib_installed:
        fig = create_demo_figure()
        pp.add_matplotlib_figure(fig, title_slide, PPTXPosition(0.3, 0.4))
        pp.add_matplotlib_figure(fig, title_slide, PPTXPosition(0.7, 0.4), zoom=0.4)

    table_data = []
    table_data.append([1, 2])  # rows can have different length
    table_data.append([4, slide2, 6])  # there is specific type needed for entries (implemented as text=f"{entry}")
    table_data.append(["", 8, 9])

    pp.add_table(slide2, table_data)
    pp.save("example.pptx")

    #
    # plotter = PI88Plotter(PI88Measurement("..\\resources\\AuSn_Creep\\1000uN 01 LC.tdm"))
    # fig = plotter.get_load_displacement_plot()
    # fig_width = fig.get_figwidth()
    # zoom = 1
    # picture = pptx.add_matplotlib_figure(fig, title_slide, PPTXPosition(0.2, 0.25),
    #                                      width=Inches(fig_width * zoom))
    #
    #
    # zoom = 0.2
    # pptx.add_matplotlib_figure(fig, title_slide, PPTXPosition(0.7, 0.25), width=Inches(fig_width * zoom))
    #
    # pptx.add_text_box(slide2, "This is the first paragraph\n... and here comes the second", PPTXPosition(0.7, 0.7))
    # pptx.save("delme_pptx_creator_demo.pptx")

    # reposition picture:
    # picture.left = Inches(1)
    # picture.top = Inches(3)


if __name__ == '__main__':
    run()
