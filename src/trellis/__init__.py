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

    # もし、柱のリストがあれば
    if 'pillars' in document and (pillars_list := document['pillars']):

        for pillar_dict in pillars_list:
            if 'baseColor' in pillar_dict and (baseColor := pillar_dict['baseColor']):
                left = pillar_dict['left']
                top = pillar_dict['top']
                width = pillar_dict['width']
                height = pillar_dict['height']

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

    # もし、柱のリストがあれば
    if 'pillars' in document and (pillars_list := document['pillars']):

        for pillar_dict in pillars_list:
            # もし、カードの辞書があれば
            if 'cards' in pillar_dict and (card_dict_list := pillar_dict['cards']):

                for card_dict in card_dict_list:
                    if 'shadowColor' in card_dict:
                        card_shadow_color = card_dict['shadowColor']

                        card_rect = get_rectangle(rectangle_dict=card_dict)

                        # 端子の影を描く
                        fill_rectangle(
                                ws=ws,
                                column_th=card_rect.left_obj.cell_th + square_unit,
                                row_th=card_rect.top_obj.cell_th + square_unit,
                                columns=card_rect.width_columns,
                                rows=card_rect.height_rows,
                                fill_obj=tone_and_color_name_to_fill_obj(card_shadow_color))


def render_all_cards(document, ws):
    """全てのカードの描画
    """
    print('全てのカードの描画')

    # もし、柱のリストがあれば
    if 'pillars' in document and (pillars_list := document['pillars']):

        for pillar_dict in pillars_list:

            # 柱と柱の隙間（隙間柱）は無視する
            if 'baseColor' not in pillar_dict or not pillar_dict['baseColor']:
                continue

            baseColor = pillar_dict['baseColor']
            card_list = pillar_dict['cards']

            for card_dict in card_list:

                card_rect = get_rectangle(rectangle_dict=card_dict)

                # ヘッダーの矩形の枠線を描きます
                draw_rectangle(
                        ws=ws,
                        column_th=card_rect.left_obj.cell_th,
                        row_th=card_rect.top_obj.cell_th,
                        columns=card_rect.width_columns,
                        rows=card_rect.height_rows)

                if 'paperStrips' in card_dict:
                    paper_strip_list = card_dict['paperStrips']

                    for index, paper_strip in enumerate(paper_strip_list):

                        # 短冊１行の描画
                        render_paper_strip(
                                ws=ws,
                                paper_strip=paper_strip,
                                column_th=card_rect.left_obj.cell_th,
                                row_th=index * square_unit + card_rect.top_obj.cell_th,
                                columns=card_rect.width_columns,
                                rows=card_rect.height_rows)


def render_all_terminal_shadows(document, ws):
    """全ての端子の影の描画
    """
    print('全ての端子の影の描画')

    # もし、柱のリストがあれば
    if 'pillars' in document and (pillars_list := document['pillars']):

        for pillar_dict in pillars_list:
            # もし、端子のリストがあれば
            if 'terminals' in pillar_dict and (terminals_list := pillar_dict['terminals']):

                for terminal_dict in terminals_list:

                    terminal_rect = get_rectangle(rectangle_dict=terminal_dict)
                    terminal_shadow_color = terminal_dict['shadowColor']

                    # 端子の影を描く
                    fill_rectangle(
                            ws=ws,
                            column_th=terminal_rect.left_obj.cell_th + square_unit,
                            row_th=terminal_rect.top_obj.cell_th + square_unit,
                            columns=9,
                            rows=9,
                            fill_obj=tone_and_color_name_to_fill_obj(terminal_shadow_color))


def render_all_terminals(document, ws):
    """全ての端子の描画
    """
    print('全ての端子の描画')

    # もし、柱のリストがあれば
    if 'pillars' in document and (pillars_list := document['pillars']):

        for pillar_dict in pillars_list:
            # もし、端子のリストがあれば
            if 'terminals' in pillar_dict and (terminals_list := pillar_dict['terminals']):

                for terminal_dict in terminals_list:

                    terminal_pixel_art = terminal_dict['pixelArt']
                    terminal_rect = get_rectangle(rectangle_dict=terminal_dict)

                    if terminal_pixel_art == 'start':
                        # 始端のドット絵を描く
                        fill_start_terminal(
                            ws=ws,
                            column_th=terminal_rect.left_obj.cell_th,
                            row_th=terminal_rect.top_obj.cell_th)
                    
                    elif terminal_pixel_art == 'end':
                        # 終端のドット絵を描く
                        fill_end_terminal(
                            ws=ws,
                            column_th=terminal_rect.left_obj.cell_th,
                            row_th=terminal_rect.top_obj.cell_th)


class Square():
    """マス
    """


    @staticmethod
    def from_main_and_sub(main_number, sub_number):
        if sub_number == 0:
            return Square(main_number)
        
        else:
            return Square(f'{main_number}o{sub_number}')


    def __init__(self, value):

        if isinstance(value, str):
            main_number, sub_number = map(int, value.split('o', 2))
            self._sub_number = sub_number
            self._main_number = main_number
        else:
            self._sub_number = 0
            self._main_number = value

        if self._sub_number == 0:
            self._var_value = self._main_number
        else:
            self._var_value = f'{self._main_number}o{self._sub_number}'

        self._cell_th = None


    @property
    def var_value(self):
        return self._var_value


    @property
    def main_number(self):
        return self._main_number


    @property
    def sub_number(self):
        return self._sub_number


    @property
    def cell_th(self):
        """1から始まるセル番号
        """
        if not self._cell_th:
            self._cell_th = self._main_number * square_unit + self._sub_number + 1

        return self._cell_th


    def offset(self, var_value):
        square = Square(var_value)
        sub_number = self._sub_number + square.sub_number
        main_number = self._main_number + square.main_number + sub_number // square_unit
        sub_number = sub_number % square_unit
        return Square.from_main_and_sub(main_number=main_number, sub_number=sub_number)


class Rectangle():
    """矩形
    """


    def __init__(self, left, sub_left, top, sub_top, width, sub_width, height, sub_height):
        """初期化
        """
        self._left_obj = Square.from_main_and_sub(main_number=left, sub_number=sub_left)
        self._width = width
        self._sub_width = sub_width

        self._left_column_th = None
        self._width_columns = None
        self._right_obj = None

        self._top_obj = Square.from_main_and_sub(main_number=top, sub_number=sub_top)
        self._height = height
        self._sub_height = sub_height

        self._top_row_th = None
        self._height_rows = None


    def _calculate_right(self):
        # サブ右＝サブ左＋サブ幅
        sum_sub_right = self._left_obj.sub_number + self._sub_width
        self._right_obj = Square.from_main_and_sub(
                main_number=self._left_obj.main_number + self._width + sum_sub_right // square_unit,
                sub_number=sum_sub_right % square_unit)


    @property
    def left_obj(self):
        return self._left_obj


    @property
    def right_obj(self):
        """矩形の右位置
        """
        if not self._right_obj:
            self._calculate_right()
        return self._right_obj


    @property
    def top_obj(self):
        return self._top_obj


    @property
    def width(self):
        return self._width


    @property
    def sub_width(self):
        return self._sub_width


    @property
    def width_columns(self):
        if not self._width_columns:
            self._width_columns = self._width * square_unit + self._sub_width
        
        return self._width_columns


    @property
    def height(self):
        return self._height


    @property
    def height_rows(self):
        if not self._height_rows:
            self._height_rows = self._height * square_unit + self._sub_height
        
        return self._height_rows


    @property
    def sub_height(self):
        return self._sub_height


def get_rectangle(rectangle_dict):
    """ラインテープのセグメントの矩形情報を取得
    """
    left = rectangle_dict['left']
    sub_left = 0
    if isinstance(left, str):
        left, sub_left = map(int, left.split('o', 2))
    
    top = rectangle_dict['top']
    sub_top = 0
    if isinstance(top, str):
        top, sub_top = map(int, top.split('o', 2))

    # right は、その数を含まない
    if 'right' in rectangle_dict:
        right = rectangle_dict['right']
        sub_right = 0
        if isinstance(right, str):
            right, sub_right = map(int, right.split('o', 2))

        width = right - left
        sub_width = sub_right - sub_left

    else:
        width = rectangle_dict['width']
        sub_width = 0
        if isinstance(width, str):
            width, sub_width = map(int, width.split('o', 2))

    # bottom は、その数を含まない
    if 'bottom' in rectangle_dict:
        bottom = rectangle_dict['bottom']
        sub_bottom = 0
        if isinstance(bottom, str):
            bottom, sub_bottom = map(int, bottom.split('o', 2))

        height = bottom - top
        sub_height = sub_bottom - sub_top

    else:
        height = rectangle_dict['height']
        sub_height = 0
        if isinstance(height, str):
            height, sub_height = map(int, height.split('o', 2))

    return Rectangle(
            left=left,
            sub_left=sub_left,
            top=top,
            sub_top=sub_top,
            width=width,
            sub_width=sub_width,
            height=height,
            sub_height=sub_height)


def render_all_line_tape_shadows(document, ws):
    """全てのラインテープの影の描画
    """
    print('全てのラインテープの影の描画')

    # もし、ラインテープの配列があれば
    if 'lineTapes' in document and (line_tape_list := document['lineTapes']):

        for line_tape_dict in line_tape_list:
            for segment_dict in line_tape_dict['segments']:
                if 'shadowColor' in segment_dict and (line_tape_shadow_color := segment_dict['shadowColor']):
                    segment_rect = get_rectangle(rectangle_dict=segment_dict)

                    # 端子の影を描く
                    fill_rectangle(
                            ws=ws,
                            column_th=segment_rect.left_obj.cell_th + square_unit,
                            row_th=segment_rect.top_obj.cell_th + square_unit,
                            columns=segment_rect.width_columns,
                            rows=segment_rect.height_rows,
                            fill_obj=tone_and_color_name_to_fill_obj(line_tape_shadow_color))


def render_all_line_tapes(document, ws):
    """全てのラインテープの描画
    """
    print('全てのラインテープの描画')

    # もし、ラインテープの配列があれば
    if 'lineTapes' in document and (line_tape_list := document['lineTapes']):

        # 各ラインテープ
        for line_tape_dict in line_tape_list:

            line_tape_outline_color = None
            if 'outlineColor' in line_tape_dict:
                line_tape_outline_color = line_tape_dict['outlineColor']

            # 各セグメント
            for segment_dict in line_tape_dict['segments']:

                line_tape_direction = None
                if 'direction' in segment_dict:
                    line_tape_direction = segment_dict['direction']

                if 'color' in segment_dict:
                    line_tape_color = segment_dict['color']

                    segment_rect = get_rectangle(rectangle_dict=segment_dict)

                    # ラインテープを描く
                    fill_obj = tone_and_color_name_to_fill_obj(line_tape_color)
                    fill_rectangle(
                            ws=ws,
                            column_th=segment_rect.left_obj.cell_th,
                            row_th=segment_rect.top_obj.cell_th,
                            columns=segment_rect.width_columns,
                            rows=segment_rect.height_rows,
                            fill_obj=fill_obj)

                    # （あれば）アウトラインを描く
                    if line_tape_outline_color and line_tape_direction:
                        outline_fill_obj = tone_and_color_name_to_fill_obj(line_tape_outline_color)

                        # （共通処理）垂直方向
                        if line_tape_direction in ['from_here.falling_down', 'after_go_right.turn_falling_down', 'after_go_left.turn_up', 'after_go_left.turn_falling_down']:
                            # 左辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.cell_th - 1,
                                    row_th=segment_rect.top_obj.cell_th + 1,
                                    columns=1,
                                    rows=segment_rect.height_rows - 2,
                                    fill_obj=outline_fill_obj)

                            # 右辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.cell_th + segment_rect.width_columns,
                                    row_th=segment_rect.top_obj.cell_th + 1,
                                    columns=1,
                                    rows=segment_rect.height_rows - 2,
                                    fill_obj=outline_fill_obj)
                        
                        # （共通処理）水平方向
                        elif line_tape_direction in ['after_falling_down.turn_right', 'continue.go_right', 'after_falling_down.turn_left', 'continue.go_left', 'after_up.turn_right', 'from_here.go_right']:
                            # 上辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.cell_th + square_unit,
                                    row_th=segment_rect.top_obj.cell_th - 1,
                                    columns=segment_rect.width_columns - 2 * square_unit,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                            # 下辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.cell_th + square_unit,
                                    row_th=segment_rect.top_obj.cell_th + segment_rect.height_rows,
                                    columns=segment_rect.width_columns - 2 * square_unit,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                        # ここから落ちていく
                        if line_tape_direction == 'from_here.falling_down':
                            # 左辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.cell_th - 1,
                                    row_th=segment_rect.top_obj.cell_th,
                                    columns=1,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                            # 右辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.cell_th + segment_rect.width_columns,
                                    row_th=segment_rect.top_obj.cell_th,
                                    columns=1,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                        # 落ちたあと、右折
                        elif line_tape_direction == 'after_falling_down.turn_right':
                            # 左辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.cell_th - 1,
                                    row_th=segment_rect.top_obj.cell_th - 1,
                                    columns=1,
                                    rows=2,
                                    fill_obj=outline_fill_obj)

                            # 下辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.cell_th - 1,
                                    row_th=segment_rect.top_obj.cell_th + 1,
                                    columns=square_unit + 1,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                        # そのまま右進
                        elif line_tape_direction == 'continue.go_right':
                            # 上辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.cell_th - square_unit,
                                    row_th=segment_rect.top_obj.cell_th - 1,
                                    columns=2 * square_unit,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                            # 下辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.cell_th - square_unit,
                                    row_th=segment_rect.top_obj.cell_th + 1,
                                    columns=2 * square_unit,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                        # 右進から落ちていく
                        elif line_tape_direction == 'after_go_right.turn_falling_down':
                            # 上辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.cell_th - square_unit,
                                    row_th=segment_rect.top_obj.cell_th - 1,
                                    columns=2 * square_unit,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                            # 下辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.cell_th - square_unit,
                                    row_th=segment_rect.top_obj.cell_th + 1,
                                    columns=square_unit,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                            # 右辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.cell_th + segment_rect.width_columns,
                                    row_th=segment_rect.top_obj.cell_th - 1,
                                    columns=1,
                                    rows=2,
                                    fill_obj=outline_fill_obj)

                        # 落ちたあと左折
                        elif line_tape_direction == 'after_falling_down.turn_left':
                            # 右辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.cell_th + segment_rect.width_columns,
                                    row_th=segment_rect.top_obj.cell_th - 1,
                                    columns=1,
                                    rows=2,
                                    fill_obj=outline_fill_obj)

                            # 下辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.cell_th + segment_rect.width_columns - square_unit,
                                    row_th=segment_rect.top_obj.cell_th + 1,
                                    columns=square_unit + 1,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                        # そのまま左進
                        elif line_tape_direction == 'continue.go_left':
                            # 上辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.cell_th + square_unit,
                                    row_th=segment_rect.top_obj.cell_th - 1,
                                    columns=segment_rect.width_columns,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                            # 下辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.cell_th + square_unit,
                                    row_th=segment_rect.top_obj.cell_th + 1,
                                    columns=segment_rect.width_columns,
                                    rows=1,
                                    fill_obj=outline_fill_obj)
                        
                        # 左進から上っていく
                        elif line_tape_direction == 'after_go_left.turn_up':
                            # 下辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.cell_th,
                                    row_th=segment_rect.top_obj.cell_th + segment_rect.height_rows,
                                    columns=2 * square_unit,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                            # 左辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.cell_th - 1,
                                    row_th=segment_rect.top_obj.cell_th + segment_rect.height_rows - 2,
                                    columns=1,
                                    rows=3,
                                    fill_obj=outline_fill_obj)
                            
                            # 右辺（横長）を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.cell_th + square_unit,
                                    row_th=segment_rect.top_obj.cell_th + segment_rect.height_rows - 2,
                                    columns=square_unit,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                        # 上がってきて右折
                        elif line_tape_direction == 'after_up.turn_right':
                            # 左辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.cell_th - 1,
                                    row_th=segment_rect.top_obj.cell_th,
                                    columns=1,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                            # 上辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.cell_th - 1,
                                    row_th=segment_rect.top_obj.cell_th - 1,
                                    columns=square_unit + 1,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                        # 左進から落ちていく
                        elif line_tape_direction == 'after_go_left.turn_falling_down':
                            # 上辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.cell_th,
                                    row_th=segment_rect.top_obj.cell_th - 1,
                                    columns=2 * square_unit,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                            # 左辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.cell_th - 1,
                                    row_th=segment_rect.top_obj.cell_th - 1,
                                    columns=1,
                                    rows=segment_rect.height_rows,
                                    fill_obj=outline_fill_obj)

                            # 右辺（横長）を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.cell_th + square_unit + 1,
                                    row_th=segment_rect.top_obj.cell_th + 1,
                                    columns=square_unit - 1,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                        # ここから右進
                        elif line_tape_direction == 'from_here.go_right':
                            # 上辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.cell_th,
                                    row_th=segment_rect.top_obj.cell_th - 1,
                                    columns=square_unit,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                            # 下辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.cell_th,
                                    row_th=segment_rect.top_obj.cell_th + 1,
                                    columns=square_unit,
                                    rows=1,
                                    fill_obj=outline_fill_obj)


def resolve_auto_shadow(document, column_th, row_th):
    """影の自動設定を解決する"""

    # もし、影の色の対応付けがあれば
    if 'shadowColorMappings' in document and (shadow_color_dict := document['shadowColorMappings']):

        # もし、柱のリストがあれば
        if 'pillars' in document and (pillars_list := document['pillars']):

            for pillar_dict in pillars_list:

                # 柱と柱の隙間（隙間柱）は無視する
                if 'baseColor' not in pillar_dict or not pillar_dict['baseColor']:
                    continue

                pillar_rect = get_rectangle(rectangle_dict=pillar_dict)
                base_color = pillar_dict['baseColor']

                # もし、矩形の中に、指定の点が含まれたなら
                if pillar_rect.left_obj.cell_th <= column_th and column_th < pillar_rect.left_obj.cell_th + pillar_rect.width_columns and \
                    pillar_rect.top_obj.cell_th <= row_th and row_th < pillar_rect.top_obj.cell_th + pillar_rect.height_rows:

                    return shadow_color_dict[base_color]

    # 該当なし
    return shadow_color_dict['paper_color']


def edit_document_and_solve_auto_shadow(document):
    """ドキュメントに対して、影の自動設定の編集を行います
    """

    # もし、柱のリストがあれば
    if 'pillars' in document and (pillars_list := document['pillars']):

        for pillar_dict in pillars_list:
            # もし、カードの辞書があれば
            if 'cards' in pillar_dict and (card_dict_list := pillar_dict['cards']):

                for card_dict in card_dict_list:
                    if 'shadowColor' in card_dict and (card_shadow_color := card_dict['shadowColor']):

                        if card_shadow_color == 'auto':
                            card_rect = get_rectangle(rectangle_dict=card_dict)

                            # 影に自動が設定されていたら、解決する
                            if solved_tone_and_color_name := resolve_auto_shadow(
                                    document=document,
                                    column_th=card_rect.left_obj.cell_th + square_unit,
                                    row_th=card_rect.top_obj.cell_th + square_unit):
                                card_dict['shadowColor'] = solved_tone_and_color_name

            # もし、端子のリストがあれば
            if 'terminals' in pillar_dict and (terminals_list := pillar_dict['terminals']):

                for terminal_dict in terminals_list:
                    if 'shadowColor' in terminal_dict and (terminal_shadow_color := terminal_dict['shadowColor']):

                        if terminal_shadow_color == 'auto':
                            terminal_rect = get_rectangle(rectangle_dict=terminal_dict)

                            # 影に自動が設定されていたら、解決する
                            if solved_tone_and_color_name := resolve_auto_shadow(
                                    document=document,
                                    column_th=terminal_rect.left_obj.cell_th + square_unit,
                                    row_th=terminal_rect.top_obj.cell_th + square_unit):
                                terminal_dict['shadowColor'] = solved_tone_and_color_name

    # もし、ラインテープのリストがあれば
    if 'lineTapes' in document and (line_tape_list := document['lineTapes']):

        for line_tape_dict in line_tape_list:
            # もし、セグメントのリストがあれば
            if 'segments' in line_tape_dict and (segment_list := line_tape_dict['segments']):

                for segment_dict in segment_list:
                    if 'shadowColor' in segment_dict and (segment_shadow_color := segment_dict['shadowColor']) and segment_shadow_color == 'auto':
                        segment_rect = get_rectangle(rectangle_dict=segment_dict)

                        # NOTE 影が指定されているということは、浮いているということでもある

                        # 影に自動が設定されていたら、解決する
                        if solved_tone_and_color_name := resolve_auto_shadow(
                                document=document,
                                column_th=segment_rect.left_obj.cell_th + square_unit,
                                row_th=segment_rect.top_obj.cell_th + square_unit):
                            segment_dict['shadowColor'] = solved_tone_and_color_name


def split_segment_by_pillar(document, line_tape_segment_list, line_tape_segment_dict):
    """柱を跨ぐとき、ラインテープを分割します
    NOTE 柱は左から並んでいるものとする
    NOTE 柱の縦幅は十分に広いものとする
    NOTE 柱にサブ位置はない
    """

    new_segment_list = []

    #print('柱を跨ぐとき、ラインテープを分割します')
    segment_rect = get_rectangle(rectangle_dict=line_tape_segment_dict)

    direction = line_tape_segment_dict['direction']

    splitting_segments = []


    # TODO とりあえず、落下後の左折だけ考える。他は後で考える
    # 左進より、右進の方がプログラムが簡単
    if direction == 'after_falling_down.turn_right':
        #print('とりあえず、落下後の左折だけ考える。他は後で考える')

        # もし、柱のリストがあれば
        if 'pillars' in document and (pillars_list := document['pillars']):
            #print(f'{len(pillars_list)=}')

            # 各柱
            for pillar_dict in pillars_list:
                pillar_rect = get_rectangle(rectangle_dict=pillar_dict)

                #print(f'（条件）ラインテープの左端と右端の内側に、柱の左端があるか判定 {segment_rect.left_obj.main_number=} <= {pillar_rect.left_obj.main_number=} <  {segment_rect.right_obj.main_number=} 判定：{segment_rect.left_obj.main_number <= pillar_rect.left_obj.main_number and pillar_rect.left_obj.main_number < segment_rect.right_obj.main_number}')
                # とりあえず、ラインテープの左端と右端の内側に、柱の左端があるか判定
                if segment_rect.left_obj.main_number < pillar_rect.left_obj.main_number and pillar_rect.left_obj.main_number < segment_rect.right_obj.main_number:
                    print(f'（判定）ラインテープの左端より右と右端の内側に、柱の左端がある')

                # NOTE テープは浮いています
                #print(f'（条件）ラインテープの左端と右端の内側に、柱の右端があるか判定 {segment_rect.left_obj.main_number=} <= {pillar_rect.right_obj.main_number=} <  {segment_rect.right_obj.main_number=} 判定：{segment_rect.left_obj.main_number <= pillar_rect.right_obj.main_number and pillar_rect.right_obj.main_number < segment_rect.right_obj.main_number}')
                # とりあえず、ラインテープの左端と右端の内側に、柱の右端があるか判定
                # FIXME Square を四則演算できるようにしたい
                if segment_rect.left_obj.main_number < pillar_rect.right_obj.main_number and pillar_rect.right_obj.main_number < segment_rect.right_obj.main_number:
                    print(f'（判定）ラインテープの（左端－１マス）より右と（右端－１マス）の内側に、柱の右端がある')

                    # 既存のセグメントを削除
                    line_tape_segment_list.remove(line_tape_segment_dict)

                    # 左側のセグメントを新規作成し、新リストに追加
                    # （計算を簡単にするため）width は使わず right を使う
                    left_segment_dict = dict(line_tape_segment_dict)
                    left_segment_dict.pop('width', None)
                    left_segment_dict['right'] = Square(pillar_rect.right_obj.var_value).offset(-1).var_value
                    left_segment_dict['color'] = 'xl_standard.xl_red'   # FIXME 動作テスト
                    new_segment_list.append(left_segment_dict)

                    # 右側のセグメントを新規作成し、既存リストに追加
                    # （計算を簡単にするため）width は使わず right を使う
                    right_segment_dict = dict(line_tape_segment_dict)
                    right_segment_dict.pop('width', None)
                    right_segment_dict['left'] = pillar_rect.right_obj.var_value
                    right_segment_dict['right'] = Square(segment_rect.right_obj.main_number).offset(-1).var_value
                    right_segment_dict['color'] = 'xl_standard.xl_violet'   # FIXME 動作テスト
                    line_tape_segment_list.append(right_segment_dict)
                    line_tape_segment_dict = right_segment_dict          # 入れ替え


    elif direction == 'after_up.turn_right':
        pass

    elif direction == 'after_falling_down.turn_left':
        pass

    
    return new_segment_list


def edit_document_and_solve_auto_split_pillar(document):
    """ドキュメントに対して、影の自動設定の編集を行います
    """
    new_splitting_segments = []

    # もし、ラインテープのリストがあれば
    if 'lineTapes' in document and (line_tape_list := document['lineTapes']):

        for line_tape_dict in line_tape_list:
            # もし、セグメントのリストがあれば
            if 'segments' in line_tape_dict and (line_tape_segment_list := line_tape_dict['segments']):

                for line_tape_segment_dict in line_tape_segment_list:
                    # もし、影があれば
                    if 'shadowColor' in line_tape_segment_dict and (shadow_color := line_tape_segment_dict['shadowColor']):
                        # 柱を跨ぐとき、ラインテープを分割します
                        new_splitting_segments.extend(split_segment_by_pillar(
                                document=document,
                                line_tape_segment_list=line_tape_segment_list,
                                line_tape_segment_dict=line_tape_segment_dict))

    # 削除用ループが終わってから追加する。そうしないと無限ループしてしまう
    for splitting_segments in new_splitting_segments:
        line_tape_segment_list.append(splitting_segments)


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
    def edit_document_and_solve_auto_shadow(document):
        global edit_document_and_solve_auto_shadow
        return edit_document_and_solve_auto_shadow(document)


    @staticmethod
    def edit_document_and_solve_auto_split_pillar(document):
        global edit_document_and_solve_auto_split_pillar
        return edit_document_and_solve_auto_split_pillar(document)


######################
# MARK: trellis_in_src
######################
trellis_in_src = TrellisInSrc()
