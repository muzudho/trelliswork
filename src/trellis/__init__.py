import openpyxl as xl
from openpyxl.styles import PatternFill, Font
from openpyxl.styles.alignment import Alignment
import json


def render_ruler(document, ws):
    """定規の描画
    """

    # Trellis では、タテ：ヨコ＝３：３ で、１ユニットセルとします。
    # また、上辺、右辺、下辺、左辺に、１セル幅の定規を置きます
    length_of_columns = document['canvas']['width'] * 3
    length_of_rows    = document['canvas']['height'] * 3

    print(f"""\
    canvas
        length_of_columns = {length_of_columns}
        length_of_rows    = {length_of_rows}
    """)

    # 行の横幅
    for column_th in range(1, length_of_columns + 1):
        column_letter = xl.utils.get_column_letter(column_th)
        ws.column_dimensions[column_letter].width = 2.7    # 2.7 characters = about 30 pixels

    # 列の高さ
    for row_th in range(1, length_of_rows + 1):
        ws.row_dimensions[row_th].height = 15    # 15 points = about 30 pixels

    # ウィンドウ枠の固定
    ws.freeze_panes = 'C2'

    # 定規の着色
    dark_gray = PatternFill(patternType='solid', fgColor='808080')
    light_gray = PatternFill(patternType='solid', fgColor='F2F2F2')
    dark_gray_font = Font(color='808080')
    light_gray_font = Font(color='F2F2F2')
    center_center_alignment = Alignment(horizontal='center', vertical='center')


    # 定規の着色　＞　上辺
    row_th = 1
    for column_th in range(4, length_of_columns - 2, 3):
        column_letter = xl.utils.get_column_letter(column_th)
        column_letter2 = xl.utils.get_column_letter(column_th + 2)
        cell = ws[f'{column_letter}{row_th}']
        
        # 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12,
        # -------- -------- -------- -----------
        # dark      light    dark     light
        #
        # - 1 する
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
    # (column_th - 1)={(column_th - 1)}
    # (column_th - 1) // 3={(column_th - 1) // 3}
    # (column_th - 1) // 3 % 2={(column_th - 1) // 3 % 2}
    # """)
        unit_cell = (column_th - 1) // 3
        is_left_end = (column_th - 1) % 3 == 0

        if is_left_end:
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

        # セル結合
        ws.merge_cells(f'{column_letter}{row_th}:{column_letter2}{row_th}')


    # 定規の着色　＞　上側の両端の１セルの隙間
    column_th_list = [
        3,                      # 定規の着色　＞　左上の１セルの隙間
        length_of_columns - 2   # 定規の着色　＞　右上の１セルの隙間
    ]
    for column_th in column_th_list:
        unit_cell = (column_th - 1) // 3
        column_letter = xl.utils.get_column_letter(column_th)
        cell = ws[f'{column_letter}{row_th}']
        if unit_cell % 2 == 0:
            cell.fill = dark_gray
        else:
            cell.fill = light_gray


    # 定規の着色　＞　左辺
    column_th = 1
    for row_th in range(1, length_of_rows - 1, 3):
        column_letter = xl.utils.get_column_letter(column_th)
        column_letter2 = xl.utils.get_column_letter(column_th + 1)
        unit_cell = (row_th - 1) // 3
        is_top_end = (row_th - 1) % 3 == 0

        cell = ws[f'{column_letter}{row_th}']
        
        if is_top_end:
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

        # セル結合
        ws.merge_cells(f'{column_letter}{row_th}:{column_letter2}{row_th + 2}')


    # 定規の着色　＞　下辺
    row_th = length_of_rows
    bottom_is_dark_gray = (row_th - 1) // 3 % 2 == 0
    for column_th in range(4, length_of_columns - 2, 3):
        column_letter = xl.utils.get_column_letter(column_th)
        column_letter2 = xl.utils.get_column_letter(column_th + 2)
        cell = ws[f'{column_letter}{row_th}']
        unit_cell = (column_th - 1) // 3
        is_left_end = (column_th - 1) % 3 == 0

        if is_left_end:
            cell.value = unit_cell
            cell.alignment = center_center_alignment
            if unit_cell % 2 == 0:
                if bottom_is_dark_gray:
                    cell.font = light_gray_font
                else:
                    cell.font = dark_gray_font
            else:
                if bottom_is_dark_gray:
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

        # セル結合
        ws.merge_cells(f'{column_letter}{row_th}:{column_letter2}{row_th}')


    # 定規の着色　＞　下側の両端の１セルの隙間
    column_th_list = [
        3,                      # 定規の着色　＞　左下の１セルの隙間
        length_of_columns - 2   # 定規の着色　＞　右下の１セルの隙間
    ]
    for column_th in column_th_list:
        unit_cell = (column_th - 1) // 3
        column_letter = xl.utils.get_column_letter(column_th)
        cell = ws[f'{column_letter}{row_th}']
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
    column_th = length_of_columns - 1
    rightest_is_dark_gray = (column_th - 1) // 3 % 2 == 0
    for row_th in range(1, length_of_rows - 1, 3):
        column_letter = xl.utils.get_column_letter(column_th)
        column_letter2 = xl.utils.get_column_letter(column_th + 1)
        unit_cell = (row_th - 1) // 3
        is_top_end = (row_th - 1) % 3 == 0

        cell = ws[f'{column_letter}{row_th}']
        
        if is_top_end:
            cell.value = unit_cell
            cell.alignment = center_center_alignment
            if unit_cell % 2 == 0:
                cell.font = light_gray_font
            else:
                cell.font = dark_gray_font

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

        # セル結合
        ws.merge_cells(f'{column_letter}{row_th}:{column_letter2}{row_th + 2}')


def render_pillar_header(document, ws):
    """柱の頭の描画
    """

    # Pillars の辞書があるはず。
    pillars_dict = document['pillars']

    print(f"""\
PILLARS
-------
""")
    for pillar_id, pillar_body in pillars_dict.items():
        header_stack_array = pillar_body['header']['stack']
        print(f"""\
{pillar_id}:
    len(header_stack_array) = {len(header_stack_array)}
""")


class TrellisInSrc():
    @staticmethod
    def render_ruler(document, ws):
        global render_ruler
        render_ruler(document, ws)


    @staticmethod
    def render_pillar_header(document, ws):
        global render_pillar_header
        render_pillar_header(document, ws)


######################
# MARK: trellis_in_src
######################
trellis_in_src = TrellisInSrc()
