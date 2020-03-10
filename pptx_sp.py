from pptx import Presentation
from pptx.util import Inches
from pptx.util import Pt
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.dml.color import RGBColor
# from pptx.enum.dml import MSO_THEME_COLOR

def add_slide(prs, layout, title):
    prs.slide_width = 11887200
    prs.slide_height = 6686550
    slide = prs.slides.add_slide(layout)
    # slide.shapes.title.text = title + "\nCurrent Incumbent: "
    shape = slide.shapes
    if layout == prs.slide_layouts[0]:
        slide.shapes.title.text = title
    if layout == prs.slide_layouts[1]:
        for shap in shape:
            if not shap.has_text_frame:
                continue
            text_frame = shap.text_frame
            text_frame.text = title
            # text_frame.vertical_anchor = MSO_ANCHOR.TOP
            p = text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
    return slide, shape


def create_table(slide, data_list, style, fontsize):
    r = 0
    for dt in data_list:
        if len(dt) > r:
            r = len(dt)
    # rows = len(data_list[0])
    c = len(data_list)
    top = Inches(1.5)
    left = Inches(0.3)
    width = Inches(12.0)
    height = Inches(0.8)

    if style == 0: # horizontally
        rows, cols = c, r
        table = slide.shapes.add_table(rows, cols, left, top, width, height).table
        for r in range(rows):  # 3
            icol = 0
            while icol < cols:
                table.cell(r, icol).text = data_list[r][icol]
                icol += 1
            print("row %d done" %r)

    elif style == 1: # vertically, col by col, 1st line - 3rd line
        rows, cols = r, c
        max_row = 5
        if r < max_row:
            diff =  c - r
            print("difference: %d" %diff)
            # if no extra row, normal
        elif r > max_row:
            diff = r - c
            print("difference: %d" %diff)
            # check the
            # if there are extra rows
            index = 0
            rowlengthmax = 0
            for dl in data_list:
                if len(dl) > rowlengthmax:
                    rowlengthmax = len(dl)
                    index = data_list.index(dl)
            # now index contains the index value of the longest list
            # extra_list = []

        table = slide.shapes.add_table(
            rows, cols+cols, left, top, width, height).table
        for r in range(cols):
            irow = 0
            while irow < rows:
                ico = r + (r+1)
                table.cell(irow, ico).text = data_list[r][irow]
                irow += 1
            print("col %d done" %r)
        for t in range(len(table.columns)):
            if t % 2 == 0:
                table.columns[t].width = Inches(1.0)
            else:
                table.columns[t].width = Inches(2.2)

    else:
        pass  # or fill in with other value
    resize_table_font(table, fontsize)


def iter_cells(table):
    for row in table.rows:
        for cell in row.cells:
            yield cell


def resize_table_font(table, size):
    for cell in iter_cells(table):
        for paragraph in cell.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(size)