from pptx import Presentation
from pptx.util import Inches
from pptx.util import Pt
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.dml.color import RGBColor
import pandas as pd
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

def create_table2(slide, df_sheet, style, fontsize):

    lst_colname = list(df_sheet.columns)
    lst_colname = [i.replace(i, str(lst_colname.index(i))) for i in lst_colname]
    df_sheet.columns = lst_colname

    srs_rc = df_sheet.count()
    srs_size = srs_rc.size
    srs_rc.reset_index(drop=True)
    srs_rc_indexmax = int(srs_rc.idxmax())

    if(srs_rc[srs_rc_indexmax] > 5):
        lst_dfcol = []
        for i in range(srs_size):
            df_col = df_sheet[str(i)]
            lst_dfcol.append(df_col)
        
        col_maxrow = lst_dfcol[srs_rc_indexmax]
        col_slice1 = col_maxrow[:3].dropna(how='all').reset_index(drop=True)
        col_slice2 = col_maxrow[3:].dropna(how='all').reset_index(drop=True)

        if(srs_rc_indexmax == 0):
            df_sheet = pd.concat([col_slice1, col_slice2, lst_dfcol[1], lst_dfcol[2]])
        elif(srs_rc_indexmax == 1):
            df_sheet = pd.concat([lst_dfcol[0], col_slice1, col_slice2,  lst_dfcol[2]])
        elif(srs_rc_indexmax == 2):
            df_sheet = pd.concat([lst_dfcol[0], lst_dfcol[1], col_slice1, col_slice2])

    df_sheet = df_sheet.dropna(how='all')
    
    
    for i, column in enumerate(df_sheet):
        colcontent = df_sheet[column] # content of the column
        li_st = [column] # column = name of column
        li_st = li_st + colcontent.values.tolist() # concat the name and the content
    

    # creation of table starts from here

    r = df_sheet.shape[0]
    c = df_sheet.shape[1]
    
    top = Inches(1.5)
    left = Inches(0.3)
    width = Inches(12.0)
    height = Inches(0.8)

    if style == 0: # horizontally
        rows = c 
        cols = r
        table = slide.shapes.add_table(rows, cols, left, top, width, height).table
        for r in range(rows):  # 3
            icol = 0
            while icol < cols:
                table.cell(r, icol).text = li_st[icol]
                icol += 1
            print("row %d done" %r)
    else: # vertically
        rows = r
        cols = c
        max_row = 5
        if rows < max_row:
            diff = r - max_row # result would be negative to indicate rownum is fine
            print("difference: row({0}) - maxrow({1}) = {2}".format(rows, max_row, diff))
            print("Number of row is less than max row\nStatus: Okay")
        else:
            diff = r - max_row # result would be positive
            print("difference: row({0}) - maxrow({1}) = {2}".format(rows, max_row, diff))
            print("number of row is bigger than max row\nStatus: Not Okay")

            




def create_table(slide, data_list, style, fontsize):
    # receives dataframe instead of list
    # convert dataframe to list here


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
        # max_row = 5
        # if r < max_row:
        #     diff =  max_row - r
        #     print("difference: maxrow({0}) - row({1}) = {2}".format(max_row, r, diff))
        #     print("Number of row is less than max row = 5\nStatus: Okay")
        #     # if no extra row, normal
        # elif r > max_row:
        #     diff = r - c
        #     print("difference: row({0}) - column({1}) = {2}".format(r, c, diff))
        #     print("number of row is bigger than max row = 5\nStatus: Not Okay")
        #     # check the
        #     # if there are extra rows
        #     index = 0
        #     rowlengthmax = 0
        #     for dl in data_list:
        #         if len(dl) > rowlengthmax:
        #             rowlengthmax = len(dl)
        #             index = data_list.index(dl) # wrong
        #     # now index contains the index value of the longest list
        #     # extra_list = []

        # fill in the table with text data from the list
        table = slide.shapes.add_table(
            rows, cols+cols, left, top, width, height).table
        for r in range(cols):
            irow = 0
            while irow < rows:
                ico = r + (r+1)
                table.cell(irow, ico).text = data_list[r][irow]
                irow += 1
            print("col %d done" %r)
        # adjust column width
        for t in range(len(table.columns)):
            if t % 2 == 0:
                table.columns[t].width = Inches(1.0)
            else:
                table.columns[t].width = Inches(2.2)

        # merge header cells
        for i in range(3):
            i = i*2
            table.cell(0,i).merge(table.cell(0,i+1))


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