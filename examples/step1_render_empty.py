"""
白紙の作成
"""

import openpyxl as xl
from openpyxl.styles import PatternFill, Font
from openpyxl.styles.alignment import Alignment
import json


# ソースファイル（JSON形式）を読込
json_file_name = './examples/data/new_paper.json'
with open(json_file_name) as f:
    document = json.load(f)


# Trellis では、タテ：ヨコ＝３：３ で、１ユニットセルとします。
# また、上辺、右辺、下辺、左辺に、１セル幅の定規を置きます
length_of_columns = document['canvas']['width'] * 3 + 2
length_of_rows    = document['canvas']['height'] * 3 + 2

print(f"""\
json_file_name = {json_file_name}
canvas
    length_of_columns = {length_of_columns}
    length_of_rows    = {length_of_rows}
""")

# ワークブックを新規生成
wb = xl.Workbook()

# ワークシート
ws = wb['Sheet']

# 行の横幅
for column_th in range(1, length_of_columns + 1):
    column_letter = xl.utils.get_column_letter(column_th)
    print(f"column_letter={column_letter}")
    ws.column_dimensions[column_letter].width = 2.7    # 2.7 characters = about 30 pixels

# 列の高さ
for row_th in range(1, length_of_rows + 1):
    ws.row_dimensions[row_th].height = 15    # 15 points = about 30 pixels

# ウィンドウ枠の固定
ws.freeze_panes = 'B2'

# 定規の着色
dark_gray = PatternFill(patternType='solid', fgColor='808080')
light_gray = PatternFill(patternType='solid', fgColor='F2F2F2')
dark_gray_font = Font(color='808080')
light_gray_font = Font(color='F2F2F2')
center_center_alignment = Alignment(horizontal='center', vertical='center')
# 定規の着色　＞　上辺
row_th = 1
for column_th in range(1, length_of_columns + 1):
    column_letter = xl.utils.get_column_letter(column_th)
    cell = ws[f'{column_letter}{row_th}']
    
    # -1, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10,
    # --------- -------- -------- ---------
    # dark      light    dark     light
    #
    # + 1 する
    #
    # 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11,
    # -------- -------- -------- ----------
    # dark     light    dark     light
    #
    # 3 で割って端数を切り捨て
    #
    # 0, 0, 0, 1, 1, 1, 2, 2, 2, 3, 3, 3,
    # -------- -------- -------- --------
    # dark     light    dark     light
    #
    # 2 で割った余り
    #
    # 0, 0, 0, 1, 1, 1, 0, 0, 0, 1, 1, 1,
    # -------- -------- -------- --------
    # dark     light    dark     light
    #
#     print(f"""\
# column_th={column_th}
# (column_th + 1)={(column_th + 1)}
# (column_th + 1) // 3={(column_th + 1) // 3}
# (column_th + 1) // 3 % 2={(column_th + 1) // 3 % 2}
# """)
    unit_cell = (column_th + 1) // 3

    is_number_display = (column_th + 1) % 3 == 1
    if is_number_display:
        cell.value = unit_cell
        cell.alignment = center_center_alignment
        if unit_cell % 2 == 0:
            cell.font = light_gray_font
        else:
            cell.font = dark_gray_font

    if unit_cell % 2 == 0:
        cell.fill = dark_gray
    else:
        cell.fill = light_gray

# 定規の着色　＞　左辺
column_th = 1
for row_th in range(1, length_of_rows + 1):
    column_letter = xl.utils.get_column_letter(column_th)
    cell = ws[f'{column_letter}{row_th}']
    
    unit_cell = (row_th + 1) // 3

    is_number_display = (row_th + 1) % 3 == 1
    if is_number_display:
        cell.value = unit_cell
        cell.alignment = center_center_alignment
        if unit_cell % 2 == 0:
            cell.font = light_gray_font
        else:
            cell.font = dark_gray_font

    if unit_cell % 2 == 0:
        cell.fill = dark_gray
    else:
        cell.fill = light_gray

# 定規の着色　＞　下辺
row_th = length_of_rows
bottom_is_dark_gray = (row_th + 1) // 3 % 2 == 0
for column_th in range(1, length_of_columns + 1):
    column_letter = xl.utils.get_column_letter(column_th)
    cell = ws[f'{column_letter}{row_th}']
    
    unit_cell = (column_th + 1) // 3

    is_number_display = (column_th + 1) % 3 == 1
    if is_number_display:
        cell.value = unit_cell
        cell.alignment = center_center_alignment
        if unit_cell % 2 == 0:
            cell.font = dark_gray_font
        else:
            cell.font = light_gray_font

    if unit_cell % 2 == 0:
        if bottom_is_dark_gray:
            cell.fill = dark_gray
        else:
            cell.fill = light_gray
    else:
        if bottom_is_dark_gray:
            cell.fill = light_gray
        else:
            cell.fill = dark_gray

# 定規の着色　＞　右辺
column_th = length_of_rows
rightest_is_dark_gray = (column_th + 1) // 3 % 2 == 0
for row_th in range(1, length_of_rows + 1):
    column_letter = xl.utils.get_column_letter(column_th)
    cell = ws[f'{column_letter}{row_th}']
    
    unit_cell = (row_th + 1) // 3

    is_number_display = (row_th + 1) % 3 == 1
    if is_number_display:
        cell.value = unit_cell
        cell.alignment = center_center_alignment
        if unit_cell % 2 == 0:
            cell.font = dark_gray_font
        else:
            cell.font = light_gray_font

    if unit_cell % 2 == 0:
        if rightest_is_dark_gray:
            cell.fill = dark_gray
        else:
            cell.fill = light_gray
    else:
        if rightest_is_dark_gray:
            cell.fill = light_gray
        else:
            cell.fill = dark_gray

# ワークブックの保存            
wb.save('./temp/examples/step1_new_paper.xlsx')
