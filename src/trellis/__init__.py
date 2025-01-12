import os
import openpyxl as xl
from openpyxl.styles import PatternFill, Font
from openpyxl.styles.alignment import Alignment
from openpyxl.styles.borders import Border, Side
from openpyxl.drawing.image import Image as XlImage
import json


# Trellis では、3x3cells で１マスとします
square_unit = 3


# エクセルの色システム（勝手に作ったったもの）
fill_palette_none = PatternFill(patternType=None)
fill_palette = {
    'xl_theme' : {
        'xl_white' : PatternFill(patternType='solid', fgColor='FFFFFF'),
        'xl_black' : PatternFill(patternType='solid', fgColor='000000'),
        'xl_red_gray' : PatternFill(patternType='solid', fgColor='E7E6E6'),
        'xl_blue_gray' : PatternFill(patternType='solid', fgColor='44546A'),
        'xl_blue' : PatternFill(patternType='solid', fgColor='5B9BD5'),
        'xl_brown' : PatternFill(patternType='solid', fgColor='ED7D31'),
        'xl_gray' : PatternFill(patternType='solid', fgColor='A5A5A5'),
        'xl_yellow' : PatternFill(patternType='solid', fgColor='FFC000'),
        'xl_naviy' : PatternFill(patternType='solid', fgColor='4472C4'),
        'xl_green' : PatternFill(patternType='solid', fgColor='70AD47'),
    },
    'xl_pale' : {
        'xl_white' : PatternFill(patternType='solid', fgColor='F2F2F2'),
        'xl_black' : PatternFill(patternType='solid', fgColor='808080'),
        'xl_red_gray' : PatternFill(patternType='solid', fgColor='AEAAAA'),
        'xl_blue_gray' : PatternFill(patternType='solid', fgColor='D6DCE4'),
        'xl_blue' : PatternFill(patternType='solid', fgColor='DDEBF7'),
        'xl_brown' : PatternFill(patternType='solid', fgColor='FCE4D6'),
        'xl_gray' : PatternFill(patternType='solid', fgColor='EDEDED'),
        'xl_yellow' : PatternFill(patternType='solid', fgColor='FFF2CC'),
        'xl_naviy' : PatternFill(patternType='solid', fgColor='D9E1F2'),
        'xl_green' : PatternFill(patternType='solid', fgColor='E2EFDA'),
    },
    'xl_light' : {
        'xl_white' : PatternFill(patternType='solid', fgColor='D9D9D9'),
        'xl_black' : PatternFill(patternType='solid', fgColor='595959'),
        'xl_red_gray' : PatternFill(patternType='solid', fgColor='757171'),
        'xl_blue_gray' : PatternFill(patternType='solid', fgColor='ACB9CA'),
        'xl_blue' : PatternFill(patternType='solid', fgColor='BDD7EE'),
        'xl_brown' : PatternFill(patternType='solid', fgColor='F8CBAD'),
        'xl_gray' : PatternFill(patternType='solid', fgColor='DBDBDB'),
        'xl_yellow' : PatternFill(patternType='solid', fgColor='FFE699'),
        'xl_naviy' : PatternFill(patternType='solid', fgColor='B4C6E7'),
        'xl_green' : PatternFill(patternType='solid', fgColor='C6E0B4'),
    },
    'xl_soft' : {
        'xl_white' : PatternFill(patternType='solid', fgColor='BFBFBF'),
        'xl_black' : PatternFill(patternType='solid', fgColor='404040'),
        'xl_red_gray' : PatternFill(patternType='solid', fgColor='3A3838'),
        'xl_blue_gray' : PatternFill(patternType='solid', fgColor='8497B0'),
        'xl_blue' : PatternFill(patternType='solid', fgColor='9BC2E6'),
        'xl_brown' : PatternFill(patternType='solid', fgColor='F4B084'),
        'xl_gray' : PatternFill(patternType='solid', fgColor='C9C9C9'),
        'xl_yellow' : PatternFill(patternType='solid', fgColor='FFD966'),
        'xl_naviy' : PatternFill(patternType='solid', fgColor='8EA9DB'),
        'xl_green' : PatternFill(patternType='solid', fgColor='A9D08E'),
    },
    'xl_strong' : {
        'xl_white' : PatternFill(patternType='solid', fgColor='A6A6A6'),
        'xl_black' : PatternFill(patternType='solid', fgColor='262626'),
        'xl_red_gray' : PatternFill(patternType='solid', fgColor='3A3838'),
        'xl_blue_gray' : PatternFill(patternType='solid', fgColor='333F4F'),
        'xl_blue' : PatternFill(patternType='solid', fgColor='2F75B5'),
        'xl_brown' : PatternFill(patternType='solid', fgColor='C65911'),
        'xl_gray' : PatternFill(patternType='solid', fgColor='7B7B7B'),
        'xl_yellow' : PatternFill(patternType='solid', fgColor='BF8F00'),
        'xl_naviy' : PatternFill(patternType='solid', fgColor='305496'),
        'xl_green' : PatternFill(patternType='solid', fgColor='548235'),
    },
    'xl_deep' : {
        'xl_white' : PatternFill(patternType='solid', fgColor='808080'),
        'xl_black' : PatternFill(patternType='solid', fgColor='0D0D0D'),
        'xl_red_gray' : PatternFill(patternType='solid', fgColor='161616'),
        'xl_blue_gray' : PatternFill(patternType='solid', fgColor='161616'),
        'xl_blue' : PatternFill(patternType='solid', fgColor='1F4E78'),
        'xl_brown' : PatternFill(patternType='solid', fgColor='833C0C'),
        'xl_gray' : PatternFill(patternType='solid', fgColor='525252'),
        'xl_yellow' : PatternFill(patternType='solid', fgColor='806000'),
        'xl_naviy' : PatternFill(patternType='solid', fgColor='203764'),
        'xl_green' : PatternFill(patternType='solid', fgColor='375623'),
    },
    'xl_standard' : {
        'xl_brown' : PatternFill(patternType='solid', fgColor='C00000'),
        'xl_red' : PatternFill(patternType='solid', fgColor='FF0000'),
        'xl_orange' : PatternFill(patternType='solid', fgColor='FFC000'),
        'xl_yellow' : PatternFill(patternType='solid', fgColor='FFFF00'),
        'xl_yellow_green' : PatternFill(patternType='solid', fgColor='92D050'),
        'xl_green' : PatternFill(patternType='solid', fgColor='00B050'),
        'xl_dodger_blue' : PatternFill(patternType='solid', fgColor='00B0F0'),
        'xl_blue' : PatternFill(patternType='solid', fgColor='0070C0'),
        'xl_naviy' : PatternFill(patternType='solid', fgColor='002060'),
        'xl_violet' : PatternFill(patternType='solid', fgColor='7030A0'),
    }
}


def tone_and_color_name_to_fill_obj(tone_and_color_name):
    """トーン名・色名を FillPattern オブジェクトに変換します
    """

    # 色が指定されていないとき、この関数を呼び出してはいけません
    if tone_and_color_name is None:
        raise Exception(f'tone_and_color_name_to_fill_obj: 色が指定されていません')

    # 背景色を［なし］にします。透明（transparent）で上書きするのと同じです
    if tone_and_color_name == 'paper_color':
        return fill_palette_none

    # ［auto］は自動で影の色を設定する機能ですが、その機能をオフにしているときは、とりあえず黒色にします
    if tone_and_color_name == 'auto':
        return fill_palette['xl_theme']['xl_black']

    tone, color = tone_and_color_name.split('.', 2)

    if tone in fill_palette:
        tone = tone.strip()
        if color in fill_palette[tone]:
            color = color.strip()
            return fill_palette[tone][color]
        
    print(f'tone_and_color_name_to_fill_obj: 色がない {tone_and_color_name=}')
    return fill_palette_none


def render_ruler(document, ws):
    """定規の描画
    """
    print("定規の描画")

    # Trellis では、タテ：ヨコ＝３：３ で、１ユニットセルとします。
    # また、上辺、右辺、下辺、左辺に、１セル幅の定規を置きます
    length_of_columns = document['canvas']['width'] * square_unit
    length_of_rows    = document['canvas']['height'] * square_unit

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
    for column_th in range(4, length_of_columns - 2, square_unit):
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
    # (column_th - 1) // square_unit={(column_th - 1) // square_unit}
    # (column_th - 1) // square_unit % 2={(column_th - 1) // square_unit % 2}
    # """)
        unit_cell = (column_th - 1) // square_unit
        is_left_end = (column_th - 1) % square_unit == 0

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
        square_unit,                            # 定規の着色　＞　左上の１セルの隙間
        length_of_columns - (square_unit - 1)   # 定規の着色　＞　右上の１セルの隙間
    ]
    for column_th in column_th_list:
        unit_cell = (column_th - 1) // square_unit
        column_letter = xl.utils.get_column_letter(column_th)
        cell = ws[f'{column_letter}{row_th}']
        if unit_cell % 2 == 0:
            cell.fill = dark_gray
        else:
            cell.fill = light_gray


    # 定規の着色　＞　左辺
    column_th = 1
    for row_th in range(1, length_of_rows - 1, square_unit):
        column_letter = xl.utils.get_column_letter(column_th)
        column_letter2 = xl.utils.get_column_letter(column_th + 1)
        unit_cell = (row_th - 1) // square_unit
        is_top_end = (row_th - 1) % square_unit == 0

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
    bottom_is_dark_gray = (row_th - 1) // square_unit % 2 == 0
    for column_th in range(4, length_of_columns - 2, square_unit):
        column_letter = xl.utils.get_column_letter(column_th)
        column_letter2 = xl.utils.get_column_letter(column_th + 2)
        cell = ws[f'{column_letter}{row_th}']
        unit_cell = (column_th - 1) // square_unit
        is_left_end = (column_th - 1) % square_unit == 0

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
        square_unit,                            # 定規の着色　＞　左下の１セルの隙間
        length_of_columns - (square_unit - 1)   # 定規の着色　＞　右下の１セルの隙間
    ]
    for column_th in column_th_list:
        unit_cell = (column_th - 1) // square_unit
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
    rightest_is_dark_gray = (column_th - 1) // square_unit % 2 == 0
    for row_th in range(1, length_of_rows - 1, square_unit):
        column_letter = xl.utils.get_column_letter(column_th)
        column_letter2 = xl.utils.get_column_letter(column_th + 1)
        unit_cell = (row_th - 1) // square_unit
        is_top_end = (row_th - 1) % square_unit == 0

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


def draw_rectangle(ws, column_th, row_th, columns, rows):
    """矩形の枠線の描画
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

    # 罫線で四角を作る　＞　左上
    cur_column_th = column_th
    column_letter = xl.utils.get_column_letter(cur_column_th)
    cur_row_th = row_th
    cell = ws[f'{column_letter}{cur_row_th}']
    cell.border = black_top_left_border

    # 罫線で四角を作る　＞　上辺
    for cur_column_th in range(column_th + 1, column_th + columns - 1):
        column_letter = xl.utils.get_column_letter(cur_column_th)
        cell = ws[f'{column_letter}{cur_row_th}']
        cell.border = black_top_border

    # 罫線で四角を作る　＞　右上
    cur_column_th = column_th + columns - 1
    column_letter = xl.utils.get_column_letter(cur_column_th)
    cell = ws[f'{column_letter}{cur_row_th}']
    cell.border = black_top_right_border

    # 罫線で四角を作る　＞　左辺
    cur_column_th = column_th
    for cur_row_th in range(row_th + 1, row_th + rows - 1):
        column_letter = xl.utils.get_column_letter(cur_column_th)
        cell = ws[f'{column_letter}{cur_row_th}']
        cell.border = black_left_border

    # 罫線で四角を作る　＞　左下
    cur_row_th = row_th + rows - 1
    cell = ws[f'{column_letter}{cur_row_th}']
    cell.border = black_bottom_left_border

    # 罫線で四角を作る　＞　下辺
    for cur_column_th in range(column_th + 1, column_th + columns - 1):
        column_letter = xl.utils.get_column_letter(cur_column_th)
        cell = ws[f'{column_letter}{cur_row_th}']
        cell.border = black_bottom_border

    # 罫線で四角を作る　＞　右下
    cur_column_th = column_th + columns - 1
    column_letter = xl.utils.get_column_letter(cur_column_th)
    cell = ws[f'{column_letter}{cur_row_th}']
    cell.border = black_bottom_right_border

    # 罫線で四角を作る　＞　右辺
    for cur_row_th in range(row_th + 1, row_th + rows - 1):
        cell = ws[f'{column_letter}{cur_row_th}']
        cell.border = black_right_border


def fill_rectangle(ws, column_th, row_th, columns, rows, fill_obj):
    """矩形を塗りつぶします
    """
    # 横へ
    for cur_column_th in range(column_th, column_th + columns):
        column_letter = xl.utils.get_column_letter(cur_column_th)

        # 縦へ
        for cur_row_th in range(row_th, row_th + rows):
            cell = ws[f'{column_letter}{cur_row_th}']
            cell.fill = fill_obj
    


def fill_pixel_art(ws, column_th, row_th, columns, rows, pixels):
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


def fill_start_terminal(ws, column_th, row_th):
    """始端を描きます
    """
    # ドット絵を描きます
    fill_pixel_art(
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


def fill_end_terminal(ws, column_th, row_th):
    """終端を描きます
    """
    # ドット絵を描きます
    fill_pixel_art(
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


def render_all_pillar_rugs(document, ws):
    """全ての柱の敷物の描画
    """
    print('全ての柱の敷物の描画')

    # もし、柱の辞書があれば
    if 'pillars' in document and (pillars_dict := document['pillars']):

        for pillar_id, whole_pillar in pillars_dict.items():
            if 'baseColor' in whole_pillar and (baseColor := whole_pillar['baseColor']):
                left = whole_pillar['left']
                top = whole_pillar['top']
                width = whole_pillar['width']
                height = whole_pillar['height']

                # 矩形を塗りつぶす
                fill_rectangle(
                        ws=ws,
                        column_th=left * square_unit + 1,
                        row_th=top * square_unit + 1,
                        columns=width * square_unit,
                        rows=height * square_unit,
                        fill_obj=tone_and_color_name_to_fill_obj(baseColor))


def render_paper_strip(ws, paper_strip, column_th, row_th, columns, rows):
    """短冊１行の描画
    """

    # 柱のヘッダーの背景色
    if 'bgColor' in paper_strip and (baseColor := paper_strip['bgColor']):
        # 矩形を塗りつぶす
        fill_rectangle(
                ws=ws,
                column_th=column_th,
                row_th=row_th,
                columns=columns,
                rows=1 * square_unit,   # １行分
                fill_obj=tone_and_color_name_to_fill_obj(baseColor))

    # インデント
    if 'indent' in paper_strip:
        indent = paper_strip['indent']
    else:
        indent = 0

    # アイコン（があれば画像をワークシートのセルに挿入）
    if 'icon' in paper_strip:
        image_basename = paper_strip['icon']  # 例： 'white-game-object.png'

        cur_column_th = column_th + (indent * square_unit)
        column_letter = xl.utils.get_column_letter(cur_column_th)
        #
        # NOTE 元の画像サイズで貼り付けられるわけではないの、何でだろう？ 60x60pixels の画像にしておくと、90x90pixels のセルに合う？
        #
        # TODO 📖 [PythonでExcelファイルに画像を挿入する/列の幅を調整する](https://qiita.com/kaba_san/items/b231a41891ebc240efc7)
        # 難しい
        #
        try:
            ws.add_image(XlImage(os.path.join('./assets/icons', image_basename)), f"{column_letter}{row_th}")
        except FileNotFoundError as e:
            print(f'FileNotFoundError {e=} {image_basename=}')


    # テキスト（があれば）
    if 'text0' in paper_strip:
        text = paper_strip['text0']
        
        # 左に１マス分のアイコンを置く前提
        icon_columns = square_unit
        cur_column_th = column_th + icon_columns + (indent * square_unit)
        column_letter = xl.utils.get_column_letter(cur_column_th)
        cell = ws[f'{column_letter}{row_th}']
        cell.value = text

    if 'text1' in paper_strip:
        text = paper_strip['text1']
        
        # 左に１マス分のアイコンを置く前提
        icon_columns = square_unit
        cur_column_th = column_th + icon_columns + (indent * square_unit)
        column_letter = xl.utils.get_column_letter(cur_column_th)
        cell = ws[f'{column_letter}{row_th + 1}']
        cell.value = text

    if 'text3' in paper_strip:
        text = paper_strip['text2']
        
        # 左に１マス分のアイコンを置く前提
        icon_columns = square_unit
        cur_column_th = column_th + icon_columns + (indent * square_unit)
        column_letter = xl.utils.get_column_letter(cur_column_th)
        cell = ws[f'{column_letter}{row_th + 2}']
        cell.value = text


def render_all_card_shadows(document, ws):
    """全てのカードの影の描画
    """
    print('全てのカードの影の描画')

    # もし、柱の辞書があれば
    if 'pillars' in document and (pillars_dict := document['pillars']):

        for pillar_id, pillar_dict in pillars_dict.items():
            # もし、カードの辞書があれば
            if 'cards' in pillar_dict and (card_dict_list := pillar_dict['cards']):

                for card_dict in card_dict_list:
                    if 'shadowColor' in card_dict:
                        card_shadow_color = card_dict['shadowColor']


                        card_left = card_dict['left']
                        card_top = card_dict['top']
                        card_width = card_dict['width']
                        card_height = card_dict['height']

                        # 端子の影を描く
                        fill_rectangle(
                                ws=ws,
                                column_th=(card_left + 1) * square_unit + 1,
                                row_th=(card_top + 1) * square_unit + 1,
                                columns=card_width * square_unit,
                                rows=card_height * square_unit,
                                fill_obj=tone_and_color_name_to_fill_obj(card_shadow_color))


def render_all_cards(document, ws):
    """全てのカードの描画
    """
    print('全てのカードの描画')

    # もし、柱の辞書があれば
    if 'pillars' in document and (pillars_dict := document['pillars']):

        for pillar_id, whole_pillar in pillars_dict.items():
            baseColor = whole_pillar['baseColor']
            card_list = whole_pillar['cards']

            for card in card_list:
                card_left = card['left']
                card_top = card['top']
                card_width = card['width']
                card_height = card['height']

                column_th = card_left * square_unit + 1
                row_th = card_top * square_unit + 1
                columns = card_width * square_unit
                rows = card_height * square_unit

                # ヘッダーの矩形の枠線を描きます
                draw_rectangle(
                        ws=ws,
                        column_th=column_th,
                        row_th=row_th,
                        columns=columns,
                        rows=rows)

                if 'paperStrips' in card:
                    paper_strip_list = card['paperStrips']

                    for paper_strip in paper_strip_list:

                        # 短冊１行の描画
                        render_paper_strip(
                                ws=ws,
                                paper_strip=paper_strip,
                                column_th=column_th,
                                row_th=row_th,
                                columns=columns,
                                rows=rows)
                        
                        row_th += square_unit


def render_all_terminal_shadows(document, ws):
    """全ての端子の影の描画
    """
    print('全ての端子の影の描画')

    # もし、柱の辞書があれば
    if 'pillars' in document and (pillars_dict := document['pillars']):

        for pillar_id, pillar_dict in pillars_dict.items():
            # もし、端子のリストがあれば
            if 'terminals' in pillar_dict and (terminals_list := pillar_dict['terminals']):

                for terminal_dict in terminals_list:
                    terminal_left = terminal_dict['left']
                    terminal_top = terminal_dict['top']
                    terminal_shadow_color = terminal_dict['shadowColor']

                    # 端子の影を描く
                    fill_rectangle(
                            ws=ws,
                            column_th=(terminal_left + 1) * square_unit + 1,
                            row_th=(terminal_top + 1) * square_unit + 1,
                            columns=9,
                            rows=9,
                            fill_obj=tone_and_color_name_to_fill_obj(terminal_shadow_color))


def render_all_terminals(document, ws):
    """全ての端子の描画
    """
    print('全ての端子の描画')

    # もし、柱の辞書があれば
    if 'pillars' in document and (pillars_dict := document['pillars']):

        for pillar_id, pillar_dict in pillars_dict.items():
            # もし、端子のリストがあれば
            if 'terminals' in pillar_dict and (terminals_list := pillar_dict['terminals']):

                for terminal_dict in terminals_list:
                    terminal_pixel_art = terminal_dict['pixelArt']
                    terminal_left = terminal_dict['left']
                    terminal_top = terminal_dict['top']

                    if terminal_pixel_art == 'start':
                        # 始端のドット絵を描く
                        fill_start_terminal(
                            ws=ws,
                            column_th=terminal_left * square_unit + 1,
                            row_th=terminal_top * square_unit + 1)
                    
                    elif terminal_pixel_art == 'end':
                        # 終端のドット絵を描く
                        fill_end_terminal(
                            ws=ws,
                            column_th=terminal_left * square_unit + 1,
                            row_th=terminal_top * square_unit + 1)


def render_all_line_tape_shadows(document, ws):
    """全てのラインテープの影の描画
    """
    print('全てのラインテープの影の描画')

    # もし、ラインテープの配列があれば
    if 'lineTapes' in document and (line_tape_list := document['lineTapes']):

        for line_tape_dict in line_tape_list:
            segments_dict = line_tape_dict['segments']

            for segment_dict in segments_dict:
                if 'shadowColor' in segment_dict and (line_tape_shadow_color := segment_dict['shadowColor']):
                    line_tape_left = segment_dict['left']
                    line_tape_sub_left = 0
                    if isinstance(line_tape_left, str):
                        line_tape_left, line_tape_sub_left = map(int, line_tape_left.split('o', 2))
                    
                    line_tape_top = segment_dict['top']
                    line_tape_sub_top = 0
                    if isinstance(line_tape_top, str):
                        line_tape_top, line_tape_sub_top = map(int, line_tape_top.split('o', 2))

                    # right は、その数を含まない
                    if 'right' in segment_dict:
                        line_tape_right = segment_dict['right']
                        line_tape_sub_right = 0
                        if isinstance(line_tape_right, str):
                            line_tape_right, line_tape_sub_right = map(int, line_tape_right.split('o', 2))

                        line_tape_width = line_tape_right - line_tape_left
                        line_tape_sub_width = line_tape_sub_right - line_tape_sub_left

                    else:
                        line_tape_width = segment_dict['width']
                        line_tape_sub_width = 0
                        if isinstance(line_tape_width, str):
                            line_tape_width, line_tape_sub_width = map(int, line_tape_width.split('o', 2))

                    # bottom は、その数を含まない
                    if 'bottom' in segment_dict:
                        line_tape_bottom = segment_dict['bottom']
                        line_tape_sub_bottom = 0
                        if isinstance(line_tape_bottom, str):
                            line_tape_bottom, line_tape_sub_bottom = map(int, line_tape_bottom.split('o', 2))

                        line_tape_height = line_tape_bottom - line_tape_top
                        line_tape_sub_height = line_tape_sub_bottom - line_tape_sub_top

                    else:
                        line_tape_height = segment_dict['height']
                        line_tape_sub_height = 0
                        if isinstance(line_tape_height, str):
                            line_tape_height, line_tape_sub_height = map(int, line_tape_height.split('o', 2))

                    # 端子の影を描く
                    fill_rectangle(
                            ws=ws,
                            column_th=(line_tape_left + 1) * square_unit + line_tape_sub_left + 1,
                            row_th=(line_tape_top + 1) * square_unit + line_tape_sub_top + 1,
                            columns=line_tape_width * square_unit + line_tape_sub_width,
                            rows=line_tape_height * square_unit + line_tape_sub_height,
                            fill_obj=tone_and_color_name_to_fill_obj(line_tape_shadow_color))


def render_all_line_tapes(document, ws):
    """全てのラインテープの描画
    """
    print('全てのラインテープの描画')

    # もし、ラインテープの配列があれば
    if 'lineTapes' in document and (line_tape_list := document['lineTapes']):

        for line_tape_dict in line_tape_list:
            line_tape_outline_color = None
            if 'outlineColor' in line_tape_dict:
                line_tape_outline_color = line_tape_dict['outlineColor']

            segments_dict = line_tape_dict['segments']

            for segment_dict in segments_dict:
                line_tape_direction = None
                if 'direction' in segment_dict:
                    line_tape_direction = segment_dict['direction']

                if 'color' in segment_dict:
                    line_tape_color = segment_dict['color']

                    line_tape_left = segment_dict['left']
                    line_tape_sub_left = 0
                    if isinstance(line_tape_left, str):
                        line_tape_left, line_tape_sub_left = map(int, line_tape_left.split('o', 2))
                    
                    line_tape_top = segment_dict['top']
                    line_tape_sub_top = 0
                    if isinstance(line_tape_top, str):
                        line_tape_top, line_tape_sub_top = map(int, line_tape_top.split('o', 2))

                    # right は、その数を含まない
                    if 'right' in segment_dict:
                        line_tape_right = segment_dict['right']
                        line_tape_sub_right = 0
                        if isinstance(line_tape_right, str):
                            line_tape_right, line_tape_sub_right = map(int, line_tape_right.split('o', 2))

                        line_tape_width = line_tape_right - line_tape_left
                        line_tape_sub_width = line_tape_sub_right - line_tape_sub_left

                    else:
                        line_tape_width = segment_dict['width']
                        line_tape_sub_width = 0
                        if isinstance(line_tape_width, str):
                            line_tape_width, line_tape_sub_width = map(int, line_tape_width.split('o', 2))

                    # bottom は、その数を含まない
                    if 'bottom' in segment_dict:
                        line_tape_bottom = segment_dict['bottom']
                        line_tape_sub_bottom = 0
                        if isinstance(line_tape_bottom, str):
                            line_tape_bottom, line_tape_sub_bottom = map(int, line_tape_bottom.split('o', 2))

                        line_tape_height = line_tape_bottom - line_tape_top
                        line_tape_sub_height = line_tape_sub_bottom - line_tape_sub_top

                    else:
                        line_tape_height = segment_dict['height']
                        line_tape_sub_height = 0
                        if isinstance(line_tape_height, str):
                            line_tape_height, line_tape_sub_height = map(int, line_tape_height.split('o', 2))

                    # ラインテープを描く
                    column_th = line_tape_left * square_unit + line_tape_sub_left + 1
                    row_th = line_tape_top * square_unit + line_tape_sub_top + 1
                    columns = line_tape_width * square_unit + line_tape_sub_width
                    rows = line_tape_height * square_unit + line_tape_sub_height
                    fill_obj = tone_and_color_name_to_fill_obj(line_tape_color)
                    fill_rectangle(
                            ws=ws,
                            column_th=column_th,
                            row_th=row_th,
                            columns=columns,
                            rows=rows,
                            fill_obj=fill_obj)

                    # （あれば）アウトラインを描く
                    if line_tape_outline_color and line_tape_direction:
                        outline_fill_obj = tone_and_color_name_to_fill_obj(line_tape_outline_color)

                        # （共通処理）垂直方向
                        if line_tape_direction in ['from_here.falling_down', 'after_go_right.turn_falling_down', 'after_go_left.turn_up', 'after_go_left.turn_falling_down']:
                            # 左辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=column_th - 1,
                                    row_th=row_th + 1,
                                    columns=1,
                                    rows=rows - 2,
                                    fill_obj=outline_fill_obj)

                            # 右辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=column_th + columns,
                                    row_th=row_th + 1,
                                    columns=1,
                                    rows=rows - 2,
                                    fill_obj=outline_fill_obj)
                        
                        # （共通処理）水平方向
                        elif line_tape_direction in ['after_falling_down.turn_right', 'continue.go_right', 'after_falling_down.turn_left', 'continue.go_left', 'after_up.turn_right', 'from_here.go_right']:
                            # 上辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=column_th + square_unit,
                                    row_th=row_th - 1,
                                    columns=columns - 2 * square_unit,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                            # 下辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=column_th + square_unit,
                                    row_th=row_th + rows,
                                    columns=columns - 2 * square_unit,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                        # ここから落ちていく
                        if line_tape_direction == 'from_here.falling_down':
                            # 左辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=column_th - 1,
                                    row_th=row_th,
                                    columns=1,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                            # 右辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=column_th + columns,
                                    row_th=row_th,
                                    columns=1,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                        # 落ちたあと、右折
                        elif line_tape_direction == 'after_falling_down.turn_right':
                            # 左辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=column_th - 1,
                                    row_th=row_th - 1,
                                    columns=1,
                                    rows=2,
                                    fill_obj=outline_fill_obj)

                            # 下辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=column_th - 1,
                                    row_th=row_th + 1,
                                    columns=square_unit + 1,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                        # そのまま右進
                        elif line_tape_direction == 'continue.go_right':
                            # 上辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=column_th - square_unit,
                                    row_th=row_th - 1,
                                    columns=2 * square_unit,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                            # 下辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=column_th - square_unit,
                                    row_th=row_th + 1,
                                    columns=2 * square_unit,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                        # 右進から落ちていく
                        elif line_tape_direction == 'after_go_right.turn_falling_down':
                            # 上辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=column_th - square_unit,
                                    row_th=row_th - 1,
                                    columns=2 * square_unit,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                            # 下辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=column_th - square_unit,
                                    row_th=row_th + 1,
                                    columns=square_unit,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                            # 右辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=column_th + columns,
                                    row_th=row_th - 1,
                                    columns=1,
                                    rows=2,
                                    fill_obj=outline_fill_obj)

                        # 落ちたあと左折
                        elif line_tape_direction == 'after_falling_down.turn_left':
                            # 右辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=column_th + columns,
                                    row_th=row_th - 1,
                                    columns=1,
                                    rows=2,
                                    fill_obj=outline_fill_obj)

                            # 下辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=column_th + columns - square_unit,
                                    row_th=row_th + 1,
                                    columns=square_unit + 1,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                        # そのまま左進
                        elif line_tape_direction == 'continue.go_left':
                            # 上辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=column_th + square_unit,
                                    row_th=row_th - 1,
                                    columns=columns,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                            # 下辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=column_th + square_unit,
                                    row_th=row_th + 1,
                                    columns=columns,
                                    rows=1,
                                    fill_obj=outline_fill_obj)
                        
                        # 左進から上っていく
                        elif line_tape_direction == 'after_go_left.turn_up':
                            # 右辺（横長）を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=column_th + square_unit,
                                    row_th=row_th + rows - 2,
                                    columns=square_unit,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                            # 下辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=column_th,
                                    row_th=row_th + rows,
                                    columns=2 * square_unit,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                            # 左辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=column_th - 1,
                                    row_th=row_th + rows - 2,
                                    columns=1,
                                    rows=3,
                                    fill_obj=outline_fill_obj)

                        # 上がってきて右折
                        elif line_tape_direction == 'after_up.turn_right':
                            # 左辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=column_th - 1,
                                    row_th=row_th,
                                    columns=1,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                            # 上辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=column_th - 1,
                                    row_th=row_th - 1,
                                    columns=square_unit + 1,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                        # 左進から落ちていく
                        elif line_tape_direction == 'after_go_left.turn_falling_down':
                            # 上辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=column_th,
                                    row_th=row_th - 1,
                                    columns=2 * square_unit,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                            # 左辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=column_th - 1,
                                    row_th=row_th - 1,
                                    columns=1,
                                    rows=rows,
                                    fill_obj=outline_fill_obj)

                            # 右辺（横長）を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=column_th + square_unit + 1,
                                    row_th=row_th + 1,
                                    columns=square_unit - 1,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                        # ここから右進
                        elif line_tape_direction == 'from_here.go_right':
                            # 上辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=column_th,
                                    row_th=row_th - 1,
                                    columns=square_unit,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                            # 下辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=column_th,
                                    row_th=row_th + 1,
                                    columns=square_unit,
                                    rows=1,
                                    fill_obj=outline_fill_obj)


def resolve_auto_shadow(document, column_th, row_th):
    """影の自動設定を解決する"""

    # もし、影の色の対応付けがあれば
    if 'shadowColorMappings' in document and (shadow_color_dict := document['shadowColorMappings']):

        # もし、柱の辞書があれば
        if 'pillars' in document and (pillars_dict := document['pillars']):

            for pillar_id, pillars_dict in pillars_dict.items():
                pillar_left = pillars_dict['left']
                pillar_top = pillars_dict['top']
                pillar_width = pillars_dict['width']
                pillar_height = pillars_dict['height']
                base_color = pillars_dict['baseColor']

                pillar_column_th = pillar_left * square_unit + 1
                pillar_row_th = pillar_top * square_unit + 1
                pillar_columns = pillar_width * square_unit
                pillar_rows = pillar_height * square_unit

                # 矩形の中に、指定の点は含まれるか？
                if pillar_column_th <= column_th and column_th <= pillar_column_th + pillar_columns and \
                    pillar_row_th <= row_th and row_th <= pillar_row_th + pillar_rows:

                    return shadow_color_dict[base_color]

    # 該当なし
    return None


def solve_auto_shadow(document):
    """ドキュメントに対して、影の自動設定の編集を行います
    """

    # もし、柱の辞書があれば
    if 'pillars' in document and (pillars_dict := document['pillars']):

        for pillar_id, pillar_dict in pillars_dict.items():
            # もし、カードの辞書があれば
            if 'cards' in pillar_dict and (card_dict_list := pillar_dict['cards']):

                for card_dict in card_dict_list:
                    if 'shadowColor' in card_dict and (card_shadow_color := card_dict['shadowColor']):

                        card_left = card_dict['left']
                        card_top = card_dict['top']
                        card_width = card_dict['width']
                        card_height = card_dict['height']

                        column_th = (card_left + 1) * square_unit + 1
                        row_th = (card_top + 1) * square_unit + 1

                        # 影に自動が設定されていたら、解決する
                        solved_tone_and_color_name = resolve_auto_shadow(document=document, column_th=column_th, row_th=row_th)
                        if solved_tone_and_color_name:
                            card_dict['shadowColor'] = solved_tone_and_color_name

    return document


class TrellisInSrc():
    @staticmethod
    def render_ruler(document, ws):
        global render_ruler
        render_ruler(document, ws)


    @staticmethod
    def render_all_terminal_shadows(document, ws):
        global render_all_terminal_shadows
        render_all_terminal_shadows(document, ws)


    @staticmethod
    def render_all_pillar_rugs(document, ws):
        global render_all_pillar_rugs
        render_all_pillar_rugs(document, ws)


    @staticmethod
    def render_all_card_shadows(document, ws):
        global render_all_card_shadows
        render_all_card_shadows(document, ws)


    @staticmethod
    def render_all_cards(document, ws):
        global render_all_cards
        render_all_cards(document, ws)


    @staticmethod
    def render_all_terminals(document, ws):
        global render_all_terminals
        render_all_terminals(document, ws)


    @staticmethod
    def render_all_line_tape_shadows(document, ws):
        global render_all_line_tape_shadows
        render_all_line_tape_shadows(document, ws)


    @staticmethod
    def render_all_line_tapes(document, ws):
        global render_all_line_tapes
        render_all_line_tapes(document, ws)


    @staticmethod
    def solve_auto_shadow(document):
        global solve_auto_shadow
        return solve_auto_shadow(document)


######################
# MARK: trellis_in_src
######################
trellis_in_src = TrellisInSrc()
