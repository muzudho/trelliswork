import os
import openpyxl as xl
from openpyxl.styles import PatternFill, Font
from openpyxl.styles.alignment import Alignment
from openpyxl.styles.borders import Border, Side
from openpyxl.drawing.image import Image as XlImage
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


def render_rectangle(ws, column_th, row_th, columns, rows, fill_obj):
    """矩形を塗りつぶします
    """
    # 横へ
    for cur_column_th in range(column_th, column_th + columns):
        column_letter = xl.utils.get_column_letter(cur_column_th)

        # 縦へ
        for cur_row_th in range(row_th, row_th + rows):
            cell = ws[f'{column_letter}{cur_row_th}']
            cell.fill = fill_obj


def render_pixel_art(ws, column_th, row_th, columns, rows, pixels):
    """ドット絵を描きます
    """
    # 背景色
    mat_black = PatternFill(patternType='solid', fgColor='080808')
    mat_white = PatternFill(patternType='solid', fgColor='E8E8E8')
    
    # 横へ
    for cur_column_th in range(column_th, column_th + columns):
        for cur_row_th in range(row_th, row_th + rows):
            column_letter = xl.utils.get_column_letter(cur_column_th)
            cell = ws[f'{column_letter}{cur_row_th}']

            pixel = pixels[cur_row_th - row_th][cur_column_th - column_th]
            if pixel == 1:
                cell.fill = mat_black
            else:
                cell.fill = mat_white


def render_start_terminal(ws, column_th, row_th):
    """始端を描きます
    """
    # ドット絵を描きます
    render_pixel_art(
            ws=ws,
            column_th=column_th,
            row_th=row_th,
            columns=9,
            rows=9,
            pixels=[
                [1, 1, 1, 1, 1, 1, 1, 1, 1],
                [1, 1, 1, 0, 0, 0, 1, 1, 1],
                [1, 0, 0, 1, 1, 1, 0, 0, 1],
                [1, 1, 0, 1, 1, 1, 1, 0, 1],
                [1, 1, 1, 0, 0, 0, 1, 1, 1],
                [1, 0, 1, 1, 1, 1, 0, 1, 1],
                [1, 0, 0, 1, 1, 1, 0, 0, 1],
                [1, 1, 1, 0, 0, 0, 1, 1, 1],
                [1, 1, 1, 1, 1, 1, 1, 1, 1],
            ])


def render_end_terminal(ws, column_th, row_th):
    """終端を描きます
    """
    # ドット絵を描きます
    render_pixel_art(
            ws=ws,
            column_th=column_th,
            row_th=row_th,
            columns=9,
            rows=9,
            pixels=[
                [1, 1, 1, 1, 1, 1, 1, 1, 1],
                [1, 0, 0, 0, 0, 0, 0, 0, 1],
                [1, 0, 1, 1, 1, 1, 1, 1, 1],
                [1, 0, 1, 1, 1, 1, 1, 1, 1],
                [1, 0, 0, 0, 0, 0, 0, 0, 1],
                [1, 0, 1, 1, 1, 1, 1, 1, 1],
                [1, 0, 1, 1, 1, 1, 1, 1, 1],
                [1, 0, 0, 0, 0, 0, 0, 0, 1],
                [1, 1, 1, 1, 1, 1, 1, 1, 1],
            ])


def render_pillar_headers(document, ws):
    """全ての柱の頭の描画
    """

    # 赤はデバッグ用
    red_side = Side(style='thick', color='FF0000')
    black_side = Side(style='thick', color='000000')

    red_top_border = Border(top=red_side)
    red_top_right_border = Border(top=red_side, right=red_side)
    red_right_border = Border(right=red_side)
    red_bottom_right_border = Border(bottom=red_side, right=red_side)
    red_bottom_border = Border(bottom=red_side)
    red_bottom_left_border = Border(bottom=red_side, left=red_side)
    red_left_border = Border(left=red_side)
    red_top_left_border = Border(top=red_side, left=red_side)

    black_top_border = Border(top=black_side)
    black_top_right_border = Border(top=black_side, right=black_side)
    black_right_border = Border(right=black_side)
    black_bottom_right_border = Border(bottom=black_side, right=black_side)
    black_bottom_border = Border(bottom=black_side)
    black_bottom_left_border = Border(bottom=black_side, left=black_side)
    black_left_border = Border(left=black_side)
    black_top_left_border = Border(top=black_side, left=black_side)

    # 柱の辞書があるはず。
    pillars_dict = document['pillars']

    # 背景色
    mat_blue = PatternFill(patternType='solid', fgColor='DDEBF7')
    mat_yellow = PatternFill(patternType='solid', fgColor='FFF2CC')

    def color_name_to_fill_obj(color_name):
        if color_name == 'blue':
            return mat_blue
        elif color_name == 'yellow':
            return mat_yellow
        else:
            return None

    print(f"""\
PILLARS
-------
""")
    for pillar_id, whole_pillar in pillars_dict.items():
        left = whole_pillar['left']
        top = whole_pillar['top']
        width = whole_pillar['width']
        height = whole_pillar['height']
        baseColor = whole_pillar['baseColor']

        # 矩形を塗りつぶす
        render_rectangle(
                ws=ws,
                column_th=left * 3 + 1,
                row_th=top * 3 + 1,
                columns=width * 3,
                rows=height * 3,
                fill_obj=color_name_to_fill_obj(baseColor))

        pillar_header = whole_pillar['header']

        header_stack_array = pillar_header['stack']
        print(f"""\
{pillar_id}:
    len(header_stack_array) = {len(header_stack_array)}
""")
        height = len(header_stack_array)

        # 罫線で四角を作る　＞　左上
        column_th = left * 3 + 1
        column_letter = xl.utils.get_column_letter(column_th)
        row_th = top * 3 + 1
        cell = ws[f'{column_letter}{row_th}']
        cell.border = black_top_left_border

        # 罫線で四角を作る　＞　上辺
        for column_th in range(left * 3 + 2, (left + width) * 3):
            column_letter = xl.utils.get_column_letter(column_th)
            cell = ws[f'{column_letter}{row_th}']
            cell.border = black_top_border

        # 罫線で四角を作る　＞　右上
        column_th = (left + width) * 3
        column_letter = xl.utils.get_column_letter(column_th)
        cell = ws[f'{column_letter}{row_th}']
        cell.border = black_top_right_border

        # 罫線で四角を作る　＞　左辺
        column_th = left * 3 + 1
        for row_th in range(top * 3 + 2, (top + height) * 3):
            column_letter = xl.utils.get_column_letter(column_th)
            cell = ws[f'{column_letter}{row_th}']
            cell.border = black_left_border

        # 罫線で四角を作る　＞　左下
        row_th = (top + height) * 3
        cell = ws[f'{column_letter}{row_th}']
        cell.border = black_bottom_left_border

        # 罫線で四角を作る　＞　下辺
        for column_th in range(left * 3 + 2, (left + width) * 3):
            column_letter = xl.utils.get_column_letter(column_th)
            cell = ws[f'{column_letter}{row_th}']
            cell.border = black_bottom_border

        # 罫線で四角を作る　＞　右下
        column_th = (left + width) * 3
        column_letter = xl.utils.get_column_letter(column_th)
        cell = ws[f'{column_letter}{row_th}']
        cell.border = black_bottom_right_border

        # 罫線で四角を作る　＞　右辺
        for row_th in range(top * 3 + 2, (top + height) * 3):
            cell = ws[f'{column_letter}{row_th}']
            cell.border = black_right_border

        row_th = top * 3 + 1
        for rectangle in header_stack_array:

            # 柱のヘッダーの背景色
            if 'bgColor' in rectangle and rectangle['bgColor']:
                # 矩形を塗りつぶす
                render_rectangle(
                        ws=ws,
                        column_th=left * 3 + 1,
                        row_th=row_th,
                        columns=width * 3,
                        rows=3,
                        fill_obj=color_name_to_fill_obj(rectangle['bgColor']))

            # インデント
            if 'indent' in rectangle:
                indent = rectangle['indent']
            else:
                indent = 0

            # アイコン（があれば画像をワークシートのセルに挿入）
            if 'icon' in rectangle:
                image_basename = rectangle['icon']  # 例： 'white-game-object.png'
                print(f'image_basename=[{image_basename}]')

                column_th = left * 3 + 3 * indent + 1
                column_letter = xl.utils.get_column_letter(column_th)
                #
                # NOTE 元の画像サイズで貼り付けられるわけではないの、何でだろう？ 60x60pixels の画像にしておくと、90x90pixels のセルに合う？
                #
                # TODO 📖 [PythonでExcelファイルに画像を挿入する/列の幅を調整する](https://qiita.com/kaba_san/items/b231a41891ebc240efc7)
                # 難しい
                #
                ws.add_image(XlImage(os.path.join('./assets/icons', image_basename)), f"{column_letter}{row_th}")

            # テキスト（があれば）
            if 'text' in rectangle:
                text = rectangle['text']
                
                column_th = (left + 1) * 3 + 3 * indent + 1
                column_letter = xl.utils.get_column_letter(column_th)
                cell = ws[f'{column_letter}{row_th + 1}']
                cell.value = rectangle['text']
                 

            row_th += 3


def render_terminals(document, ws):
    """全ての端子の描画
    """
    # 柱の辞書があるはず。
    pillars_dict = document['pillars']

    for pillar_id, pillar_dict in pillars_dict.items():

        # もし、端子の辞書があれば
        if 'terminals' in pillar_dict:
            terminals_dict = pillar_dict['terminals']

            for terminal_id, terminal_dict in terminals_dict.items():
                terminal_left = terminal_dict['left']
                terminal_top = terminal_dict['top']

                if terminal_id == 'start':
                    # 始端のドット絵を描く
                    render_start_terminal(
                        ws=ws,
                        column_th=terminal_left * 3,
                        row_th=terminal_top * 3)
                elif terminal_id == 'end':
                    # 終端のドット絵を描く
                    render_end_terminal(
                        ws=ws,
                        column_th=terminal_left * 3,
                        row_th=terminal_top * 3)


class TrellisInSrc():
    @staticmethod
    def render_ruler(document, ws):
        global render_ruler
        render_ruler(document, ws)


    @staticmethod
    def render_pillar_headers(document, ws):
        global render_pillar_headers
        render_pillar_headers(document, ws)


    @staticmethod
    def render_terminals(document, ws):
        global render_terminals
        render_terminals(document, ws)


######################
# MARK: trellis_in_src
######################
trellis_in_src = TrellisInSrc()
