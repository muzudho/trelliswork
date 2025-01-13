import os
import openpyxl as xl
from openpyxl.styles import PatternFill, Font
from openpyxl.styles.alignment import Alignment
from openpyxl.styles.borders import Border, Side
from openpyxl.drawing.image import Image as XlImage
import json


# 3 ということが言いたいだけの、長い定数名。
# Trellis では、3x3cells で［大グリッド１マス分］とします
OUT_COUNTS_THAT_CHANGE_INNING = 3


######################
# MARK: InningsPitched
######################
class InningsPitched():
    """投球回。
    トレリスでは、セル番号を指定するのに使っている
    """


    @staticmethod
    def from_integer_and_decimal_part(integer_part, decimal_part):
        """整数部と小数部を指定
        """
        if decimal_part == 0:
            return InningsPitched(integer_part)
        
        else:
            return InningsPitched(f'{integer_part}o{decimal_part}')


    def __init__(self, value):

        if isinstance(value, str):
            integer_part, decimal_part = map(int, value.split('o', 2))
            self._decimal_part = decimal_part
            self._integer_part = integer_part
        else:
            self._decimal_part = 0
            self._integer_part = value

        if self._decimal_part == 0:
            self._var_value = self._integer_part
        else:
            self._var_value = f'{self._integer_part}o{self._decimal_part}'

        self._total_of_out_counts_qty = self._integer_part * OUT_COUNTS_THAT_CHANGE_INNING + self._decimal_part


    @property
    def var_value(self):
        """投球回の整数だったり、"3o2" 形式の文字列だったりします
        """
        return self._var_value


    @property
    def integer_part(self):
        """投球回の整数部"""
        return self._integer_part


    @property
    def decimal_part(self):
        """投球回の小数部"""
        return self._decimal_part


    @property
    def total_of_out_counts_qty(self):
        """0から始まるアウト・カウントの総数
        """
        return self._total_of_out_counts_qty


    @property
    def total_of_out_counts_th(self):
        """1から始まるアウト・カウントの総数
        """
        return self._total_of_out_counts_qty + 1


    def offset(self, var_value):
        """この投球回に、引数を加算した数を算出して返します"""
        l = self                       # Left operand
        r = InningsPitched(var_value)  # Right operand
        sum_decimal_part = l.decimal_part + r.decimal_part
        integer_part = l.integer_part + r.integer_part + sum_decimal_part // OUT_COUNTS_THAT_CHANGE_INNING
        return InningsPitched.from_integer_and_decimal_part(
                integer_part=integer_part,
                decimal_part=sum_decimal_part % OUT_COUNTS_THAT_CHANGE_INNING)


#################
# MARK: Rectangle
#################
class Rectangle():
    """矩形
    """


    @staticmethod
    def from_dict(rectangle_dict):
        """ラインテープのセグメントの矩形情報を取得
        """
        main_left = rectangle_dict['left']
        sub_left = 0
        if isinstance(main_left, str):
            main_left, sub_left = map(int, main_left.split('o', 2))
        
        main_top = rectangle_dict['top']
        sub_top = 0
        if isinstance(main_top, str):
            main_top, sub_top = map(int, main_top.split('o', 2))

        # right は、その数を含まない。
        # right が指定されていれば、 width より優先する
        if 'right' in rectangle_dict:
            right = rectangle_dict['right']
            sub_right = 0
            if isinstance(right, str):
                right, sub_right = map(int, right.split('o', 2))

            main_width = right - main_left
            sub_width = sub_right - sub_left

        else:
            main_width = rectangle_dict['width']
            sub_width = 0
            if isinstance(main_width, str):
                main_width, sub_width = map(int, main_width.split('o', 2))

        # bottom は、その数を含まない。
        # bottom が指定されていれば、 width より優先する
        if 'bottom' in rectangle_dict:
            bottom = rectangle_dict['bottom']
            sub_bottom = 0
            if isinstance(bottom, str):
                bottom, sub_bottom = map(int, bottom.split('o', 2))

            main_height = bottom - main_top
            sub_height = sub_bottom - sub_top

        else:
            main_height = rectangle_dict['height']
            sub_height = 0
            if isinstance(main_height, str):
                main_height, sub_height = map(int, main_height.split('o', 2))

        return Rectangle(
                main_left=main_left,
                sub_left=sub_left,
                main_top=main_top,
                sub_top=sub_top,
                main_width=main_width,
                sub_width=sub_width,
                main_height=main_height,
                sub_height=sub_height)


    def __init__(self, main_left, sub_left, main_top, sub_top, main_width, sub_width, main_height, sub_height):
        """初期化
        """
        self._left_obj = InningsPitched.from_integer_and_decimal_part(integer_part=main_left, decimal_part=sub_left)
        self._top_obj = InningsPitched.from_integer_and_decimal_part(integer_part=main_top, decimal_part=sub_top)
        self._width_obj = InningsPitched.from_integer_and_decimal_part(integer_part=main_width, decimal_part=sub_width)
        self._height_obj = InningsPitched.from_integer_and_decimal_part(integer_part=main_height, decimal_part=sub_height)
        self._right_obj = None


    def _calculate_right(self):
        # サブ右＝サブ左＋サブ幅
        sum_sub_right = self._left_obj.decimal_part + self._width_obj.decimal_part
        self._right_obj = InningsPitched.from_integer_and_decimal_part(
                integer_part=self._left_obj.integer_part + self._width_obj.integer_part + sum_sub_right // OUT_COUNTS_THAT_CHANGE_INNING,
                decimal_part=sum_sub_right % OUT_COUNTS_THAT_CHANGE_INNING)


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
    def width_obj(self):
        return self._width_obj


    @property
    def height_obj(self):
        return self._height_obj


####################
# MARK: Color system
####################
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


#############
# MARK: Ruler
#############
def render_ruler(document, ws):
    """定規の描画
    """
    print("🔧　定規の描画")

    # Trellis では、タテ：ヨコ＝３：３ で、１ユニットセルとします。
    # また、上辺、右辺、下辺、左辺に、１セル幅の定規を置きます
    canvas_rect = Rectangle.from_dict(document['canvas'])

    # 行の横幅
    print(f"""{canvas_rect.width_obj.total_of_out_counts_th=} canvas_dict={document['canvas']}""")
    for column_th in range(1, canvas_rect.width_obj.total_of_out_counts_th):
        column_letter = xl.utils.get_column_letter(column_th)
        ws.column_dimensions[column_letter].width = 2.7    # 2.7 characters = about 30 pixels

    # 列の高さ
    for row_th in range(1, canvas_rect.height_obj.total_of_out_counts_th):
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
    #
    #   横幅が３で割り切れるとき、１投球回は 4th から始まる。２投球回を最終表示にするためには、横幅を 3 シュリンクする
    #   ■■□[  1 ][  2 ]□■■
    #   ■■                ■■
    #
    #   横幅が３で割ると１余るとき、１投球回は 4th から始まる。２投球回を最終表示にするためには、横幅を 4 シュリンクする
    #   ■■□[  1 ][  2 ]□□■■
    #   ■■                  ■■
    #
    #   横幅が３で割ると２余るとき、１投球回は 4th から始まる。２投球回を最終表示にするためには、横幅を 2 シュリンクする
    #   ■■□[  1 ][  2 ][  3 ]■■
    #   ■■                    ■■
    #
    row_th = 1
    left_th = 4
    horizontal_remain = canvas_rect.width_obj.total_of_out_counts_qty % OUT_COUNTS_THAT_CHANGE_INNING
    if horizontal_remain == 0:       
        shrink = OUT_COUNTS_THAT_CHANGE_INNING
    elif horizontal_remain == 1:
        shrink = OUT_COUNTS_THAT_CHANGE_INNING + 1
    else:
        shrink = OUT_COUNTS_THAT_CHANGE_INNING - 1

    for column_th in range(left_th, canvas_rect.width_obj.total_of_out_counts_th - shrink, OUT_COUNTS_THAT_CHANGE_INNING):
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
        unit_cell = (column_th - 1) // OUT_COUNTS_THAT_CHANGE_INNING
        is_left_end = (column_th - 1) % OUT_COUNTS_THAT_CHANGE_INNING == 0

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


    # 定規の着色　＞　左辺
    #
    #   縦幅が３で割り切れるとき、１投球回は 1th から始まる。最後の投球回は、端数なしで表示できる
    #   [  0 ][  1 ][  2 ][  3 ]
    #   ■                    ■
    #
    #   縦幅が３で割ると１余るとき、１投球回は 1th から始まる。最後の投球回は、端数１になる
    #   [  0 ][  1 ][  2 ][  3 ]□
    #   ■                      ■
    #
    #   縦幅が３で割ると２余るとき、１投球回は 1th から始まる。最後の投球回は、端数２になる
    #   [  0 ][  1 ][  2 ][  3 ]□□
    #   ■                        ■
    #
    column_th = 1
    top_th = 1
    shrink = canvas_rect.height_obj.total_of_out_counts_qty % OUT_COUNTS_THAT_CHANGE_INNING

    # # 縦の最後の要素
    # vertical_remain = canvas_rect.height_obj.total_of_out_counts_qty % OUT_COUNTS_THAT_CHANGE_INNING
    # if vertical_remain == 1:
    #     shrink += OUT_COUNTS_THAT_CHANGE_INNING

    for row_th in range(top_th, canvas_rect.height_obj.total_of_out_counts_th - shrink, OUT_COUNTS_THAT_CHANGE_INNING):
        column_letter = xl.utils.get_column_letter(column_th)
        column_letter2 = xl.utils.get_column_letter(column_th + 1)
        unit_cell = (row_th - 1) // OUT_COUNTS_THAT_CHANGE_INNING
        is_top_end = (row_th - 1) % OUT_COUNTS_THAT_CHANGE_INNING == 0

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


    # 左辺の最後の要素が端数のとき、左辺の最後の要素の左上へ着色
    #
    #       最後の端数の要素に色を塗ってもらいたいから、もう１要素着色しておく
    #
    vertical_remain = canvas_rect.height_obj.total_of_out_counts_qty % OUT_COUNTS_THAT_CHANGE_INNING
    #print(f'左辺 h_qty={canvas_rect.height_obj.total_of_out_counts_qty} {shrink=} {vertical_remain=}')
    if vertical_remain != 0:
        row_th = canvas_rect.height_obj.total_of_out_counts_th - vertical_remain
        unit_cell = (row_th - 1) // OUT_COUNTS_THAT_CHANGE_INNING
        #print(f"""左辺の最後の要素の左上へ着色 {row_th=} {unit_cell=}""")
        cell = ws[f'{column_letter}{row_th}']

        # 数字も振りたい
        if vertical_remain == 2:
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
    row_th = canvas_rect.height_obj.total_of_out_counts_th - 1
    bottom_is_dark_gray = (row_th - 1) // OUT_COUNTS_THAT_CHANGE_INNING % 2 == 0
    left_th = 4
    horizontal_remain = canvas_rect.width_obj.total_of_out_counts_qty % OUT_COUNTS_THAT_CHANGE_INNING
    if horizontal_remain == 0:       
        shrink = OUT_COUNTS_THAT_CHANGE_INNING
    elif horizontal_remain == 1:
        shrink = OUT_COUNTS_THAT_CHANGE_INNING + 1
    else:
        shrink = OUT_COUNTS_THAT_CHANGE_INNING - 1

    for column_th in range(left_th, canvas_rect.width_obj.total_of_out_counts_th - shrink, OUT_COUNTS_THAT_CHANGE_INNING):
        column_letter = xl.utils.get_column_letter(column_th)
        column_letter2 = xl.utils.get_column_letter(column_th + 2)
        cell = ws[f'{column_letter}{row_th}']
        unit_cell = (column_th - 1) // OUT_COUNTS_THAT_CHANGE_INNING
        is_left_end = (column_th - 1) % OUT_COUNTS_THAT_CHANGE_INNING == 0

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


    # 定規の着色　＞　右辺
    column_th = canvas_rect.width_obj.total_of_out_counts_th - 2
    rightest_is_dark_gray = (column_th - 1) // OUT_COUNTS_THAT_CHANGE_INNING % 2 == 0
    top_th = 1
    shrink = canvas_rect.height_obj.total_of_out_counts_qty % OUT_COUNTS_THAT_CHANGE_INNING

    for row_th in range(top_th, canvas_rect.height_obj.total_of_out_counts_th - shrink, OUT_COUNTS_THAT_CHANGE_INNING):
        column_letter = xl.utils.get_column_letter(column_th)
        column_letter2 = xl.utils.get_column_letter(column_th + 1)
        unit_cell = (row_th - 1) // OUT_COUNTS_THAT_CHANGE_INNING
        is_top_end = (row_th - 1) % OUT_COUNTS_THAT_CHANGE_INNING == 0

        cell = ws[f'{column_letter}{row_th}']
        
        if is_top_end:
            cell.value = unit_cell
            cell.alignment = center_center_alignment
            if unit_cell % 2 == 0:
                if rightest_is_dark_gray:
                    cell.font = light_gray_font
                else:
                    cell.font = dark_gray_font
            else:
                if rightest_is_dark_gray:
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

    # 右辺の最後の要素が端数のとき、右辺の最後の要素の左上へ着色
    #
    #       最後の端数の要素に色を塗ってもらいたいから、もう１要素着色しておく
    #
    vertical_remain = canvas_rect.height_obj.total_of_out_counts_qty % OUT_COUNTS_THAT_CHANGE_INNING
    #print(f'右辺 h_qty={canvas_rect.height_obj.total_of_out_counts_qty} {shrink=} {vertical_remain=}')
    if vertical_remain != 0:
        row_th = canvas_rect.height_obj.total_of_out_counts_th - vertical_remain
        unit_cell = (row_th - 1) // OUT_COUNTS_THAT_CHANGE_INNING
        #print(f"""右辺の最後の要素の左上へ着色 {row_th=} {unit_cell=}""")
        cell = ws[f'{column_letter}{row_th}']

        # 数字も振りたい
        if vertical_remain == 2:
            cell.value = unit_cell
            cell.alignment = center_center_alignment
            if unit_cell % 2 == 0:
                if rightest_is_dark_gray:
                    cell.font = light_gray_font
                else:
                    cell.font = dark_gray_font
            else:
                if rightest_is_dark_gray:
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


    # NOTE 端数の処理

    # 定規の着色　＞　左上の１セルの隙間
    row_th = 1
    column_th = OUT_COUNTS_THAT_CHANGE_INNING
    unit_cell = (column_th - 1) // OUT_COUNTS_THAT_CHANGE_INNING
    column_letter = xl.utils.get_column_letter(column_th)
    cell = ws[f'{column_letter}{row_th}']
    if unit_cell % 2 == 0:
        cell.fill = dark_gray
    else:
        cell.fill = light_gray

    # 定規の着色　＞　右上の１セルの隙間    
    row_th = 1
    side_frame_width = 2
    remain = (canvas_rect.width_obj.total_of_out_counts_qty - side_frame_width) % OUT_COUNTS_THAT_CHANGE_INNING
    column_th = canvas_rect.width_obj.total_of_out_counts_th - side_frame_width - remain
    unit_cell = (column_th - 1) // OUT_COUNTS_THAT_CHANGE_INNING
    column_letter = xl.utils.get_column_letter(column_th)
    cell = ws[f'{column_letter}{row_th}']
    if unit_cell % 2 == 0:
        cell.fill = dark_gray
    else:
        cell.fill = light_gray

    # 定規の着色　＞　左下の１セルの隙間
    row_th = canvas_rect.height_obj.total_of_out_counts_th - 1
    column_th = OUT_COUNTS_THAT_CHANGE_INNING
    unit_cell = (column_th - 1) // OUT_COUNTS_THAT_CHANGE_INNING
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

    # 定規の着色　＞　右下の１セルの隙間
    row_th = canvas_rect.height_obj.total_of_out_counts_th - 1
    side_frame_width = 2
    remain = (canvas_rect.width_obj.total_of_out_counts_qty - side_frame_width) % OUT_COUNTS_THAT_CHANGE_INNING
    column_th = canvas_rect.width_obj.total_of_out_counts_th - side_frame_width - remain
    unit_cell = (column_th - 1) // OUT_COUNTS_THAT_CHANGE_INNING
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


    # NOTE セル結合すると read only セルになるから、セル結合は、セルを編集が終わったあとで行う

    # 定規のセル結合　＞　上辺
    #
    #   横幅が３で割り切れるとき、１投球回は 4th から始まる。２投球回を最終表示にするためには、横幅を 3 シュリンクする
    #   ■■□[  1 ][  2 ]□■■
    #   ■■                ■■
    #
    #   横幅が３で割ると１余るとき、１投球回は 4th から始まる。２投球回を最終表示にするためには、横幅を 4 シュリンクする
    #   ■■□[  1 ][  2 ]□□■■
    #   ■■                  ■■
    #
    #   横幅が３で割ると２余るとき、１投球回は 4th から始まる。２投球回を最終表示にするためには、横幅を 2 シュリンクする
    #   ■■□[  1 ][  2 ][  3 ]■■
    #   ■■                    ■■
    #
    row_th = 1
    left_th = 4
    horizontal_remain = canvas_rect.width_obj.total_of_out_counts_qty % OUT_COUNTS_THAT_CHANGE_INNING
    if horizontal_remain == 0:       
        shrink = OUT_COUNTS_THAT_CHANGE_INNING
    elif horizontal_remain == 1:
        shrink = OUT_COUNTS_THAT_CHANGE_INNING + 1
    else:
        shrink = OUT_COUNTS_THAT_CHANGE_INNING - 1

    for column_th in range(left_th, canvas_rect.width_obj.total_of_out_counts_th - shrink, OUT_COUNTS_THAT_CHANGE_INNING):
        column_letter = xl.utils.get_column_letter(column_th)
        column_letter2 = xl.utils.get_column_letter(column_th + 2)
        ws.merge_cells(f'{column_letter}{row_th}:{column_letter2}{row_th}')


    # 定規のセル結合　＞　左辺
    #
    #   縦幅が３で割り切れるとき、１投球回は 1th から始まる。最後の投球回は、端数なしで表示できる
    #   [  0 ][  1 ][  2 ][  3 ]
    #   ■                    ■
    #
    #   縦幅が３で割ると１余るとき、１投球回は 1th から始まる。最後の投球回は、端数１になる
    #   [  0 ][  1 ][  2 ][  3 ]□
    #   ■                      ■
    #
    #   縦幅が３で割ると２余るとき、１投球回は 1th から始まる。最後の投球回は、端数２になる
    #   [  0 ][  1 ][  2 ][  3 ]□□
    #   ■                        ■
    #
    column_th = 1
    top_th = 1
    for row_th in range(top_th, canvas_rect.height_obj.total_of_out_counts_th - OUT_COUNTS_THAT_CHANGE_INNING, OUT_COUNTS_THAT_CHANGE_INNING):
        column_letter = xl.utils.get_column_letter(column_th)
        column_letter2 = xl.utils.get_column_letter(column_th + 1)
        ws.merge_cells(f'{column_letter}{row_th}:{column_letter2}{row_th + 2}')
    # 最後の要素
    remain = canvas_rect.height_obj.total_of_out_counts_qty % OUT_COUNTS_THAT_CHANGE_INNING
    column_letter = xl.utils.get_column_letter(column_th)
    column_letter2 = xl.utils.get_column_letter(column_th + 1)
    if remain == 0:
        row_th = canvas_rect.height_obj.integer_part * OUT_COUNTS_THAT_CHANGE_INNING + 1 - OUT_COUNTS_THAT_CHANGE_INNING
        #print(f'マージセルA h_qty={canvas_rect.height_obj.total_of_out_counts_qty} {row_th=} {remain=}')
        ws.merge_cells(f'{column_letter}{row_th}:{column_letter2}{row_th + 2}')
    elif remain == 1:
        row_th = canvas_rect.height_obj.integer_part * OUT_COUNTS_THAT_CHANGE_INNING + 1
        #print(f'マージセルH {row_th=} {remain=} {column_letter=} {column_letter2=} {canvas_rect.height_obj.integer_part=}')
        ws.merge_cells(f'{column_letter}{row_th}:{column_letter2}{row_th}')
    elif remain == 2:
        row_th = canvas_rect.height_obj.integer_part * OUT_COUNTS_THAT_CHANGE_INNING + 1
        #print(f'マージセルB h_qty={canvas_rect.height_obj.total_of_out_counts_qty} {row_th=} {remain=}')
        ws.merge_cells(f'{column_letter}{row_th}:{column_letter2}{row_th + 1}')


    # 定規のセル結合　＞　下辺
    row_th = canvas_rect.height_obj.total_of_out_counts_th - 1
    left_th = 4
    horizontal_remain = canvas_rect.width_obj.total_of_out_counts_qty % OUT_COUNTS_THAT_CHANGE_INNING
    if horizontal_remain == 0:       
        shrink = OUT_COUNTS_THAT_CHANGE_INNING
    elif horizontal_remain == 1:
        shrink = OUT_COUNTS_THAT_CHANGE_INNING + 1
    else:
        shrink = OUT_COUNTS_THAT_CHANGE_INNING - 1

    for column_th in range(left_th, canvas_rect.width_obj.total_of_out_counts_th - shrink, OUT_COUNTS_THAT_CHANGE_INNING):
        column_letter = xl.utils.get_column_letter(column_th)
        column_letter2 = xl.utils.get_column_letter(column_th + 2)
        ws.merge_cells(f'{column_letter}{row_th}:{column_letter2}{row_th}')


    # 定規のセル結合　＞　右辺
    column_th = canvas_rect.width_obj.total_of_out_counts_th - 2
    top_th = 1
    for row_th in range(top_th, canvas_rect.height_obj.total_of_out_counts_th - OUT_COUNTS_THAT_CHANGE_INNING, OUT_COUNTS_THAT_CHANGE_INNING):
        column_letter = xl.utils.get_column_letter(column_th)
        column_letter2 = xl.utils.get_column_letter(column_th + 1)
        ws.merge_cells(f'{column_letter}{row_th}:{column_letter2}{row_th + 2}')
    # 最後の要素
    remain = canvas_rect.height_obj.total_of_out_counts_qty % OUT_COUNTS_THAT_CHANGE_INNING
    column_letter = xl.utils.get_column_letter(column_th)
    column_letter2 = xl.utils.get_column_letter(column_th + 1)
    if remain == 0:
        row_th = canvas_rect.height_obj.integer_part * OUT_COUNTS_THAT_CHANGE_INNING + 1 - OUT_COUNTS_THAT_CHANGE_INNING
        #print(f'マージセルC h_qty={canvas_rect.height_obj.total_of_out_counts_qty} {row_th=} {remain=}')
        ws.merge_cells(f'{column_letter}{row_th}:{column_letter2}{row_th + 2}')
    elif remain == 1:
        row_th = canvas_rect.height_obj.integer_part * OUT_COUNTS_THAT_CHANGE_INNING + 1
        ws.merge_cells(f'{column_letter}{row_th}:{column_letter2}{row_th}')
    elif remain == 2:
        row_th = canvas_rect.height_obj.integer_part * OUT_COUNTS_THAT_CHANGE_INNING + 1
        #print(f'マージセルD h_qty={canvas_rect.height_obj.total_of_out_counts_qty} {row_th=} {remain=}')
        ws.merge_cells(f'{column_letter}{row_th}:{column_letter2}{row_th + 1}')


    # 上側の水平ルーラーの右端の端数のセル結合
    spacing = (canvas_rect.width_obj.total_of_out_counts_qty - side_frame_width) % OUT_COUNTS_THAT_CHANGE_INNING
    if spacing == 2:
        column_th = canvas_rect.width_obj.total_of_out_counts_th - side_frame_width - spacing
        row_th = 1
        column_letter = xl.utils.get_column_letter(column_th)
        column_letter2 = xl.utils.get_column_letter(column_th + spacing - 1)
        #print(f"""マージセルE {column_th=} {row_th=} {column_letter=} {column_letter2=}""")
        ws.merge_cells(f'{column_letter}{row_th}:{column_letter2}{row_th}')

    # 下側の水平ルーラーの右端の端数のセル結合
    spacing = (canvas_rect.width_obj.total_of_out_counts_qty - side_frame_width) % OUT_COUNTS_THAT_CHANGE_INNING
    if spacing == 2:
        column_th = canvas_rect.width_obj.total_of_out_counts_th - side_frame_width - spacing
        row_th = canvas_rect.height_obj.total_of_out_counts_th - 1
        column_letter = xl.utils.get_column_letter(column_th)
        column_letter2 = xl.utils.get_column_letter(column_th + spacing - 1)
        #print(f"""マージセルF {column_th=} {row_th=} {column_letter=} {column_letter2=}""")
        ws.merge_cells(f'{column_letter}{row_th}:{column_letter2}{row_th}')


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
    print('🔧　全ての柱の敷物の描画')

    # もし、柱のリストがあれば
    if 'pillars' in document and (pillars_list := document['pillars']):

        for pillar_dict in pillars_list:
            if 'baseColor' in pillar_dict and (baseColor := pillar_dict['baseColor']):
                pillar_rect = Rectangle.from_dict(pillar_dict)

                # 矩形を塗りつぶす
                fill_rectangle(
                        ws=ws,
                        column_th=pillar_rect.left_obj.total_of_out_counts_th,
                        row_th=pillar_rect.top_obj.total_of_out_counts_th,
                        columns=pillar_rect.width_obj.total_of_out_counts_qty,
                        rows=pillar_rect.height_obj.total_of_out_counts_qty,
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
                rows=1 * OUT_COUNTS_THAT_CHANGE_INNING,   # １行分
                fill_obj=tone_and_color_name_to_fill_obj(baseColor))

    # インデント
    if 'indent' in paper_strip:
        indent = paper_strip['indent']
    else:
        indent = 0

    # アイコン（があれば画像をワークシートのセルに挿入）
    if 'icon' in paper_strip:
        image_basename = paper_strip['icon']  # 例： 'white-game-object.png'

        cur_column_th = column_th + (indent * OUT_COUNTS_THAT_CHANGE_INNING)
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
        icon_columns = OUT_COUNTS_THAT_CHANGE_INNING
        cur_column_th = column_th + icon_columns + (indent * OUT_COUNTS_THAT_CHANGE_INNING)
        column_letter = xl.utils.get_column_letter(cur_column_th)
        cell = ws[f'{column_letter}{row_th}']
        cell.value = text

    if 'text1' in paper_strip:
        text = paper_strip['text1']
        
        # 左に１マス分のアイコンを置く前提
        icon_columns = OUT_COUNTS_THAT_CHANGE_INNING
        cur_column_th = column_th + icon_columns + (indent * OUT_COUNTS_THAT_CHANGE_INNING)
        column_letter = xl.utils.get_column_letter(cur_column_th)
        cell = ws[f'{column_letter}{row_th + 1}']
        cell.value = text

    if 'text3' in paper_strip:
        text = paper_strip['text2']
        
        # 左に１マス分のアイコンを置く前提
        icon_columns = OUT_COUNTS_THAT_CHANGE_INNING
        cur_column_th = column_th + icon_columns + (indent * OUT_COUNTS_THAT_CHANGE_INNING)
        column_letter = xl.utils.get_column_letter(cur_column_th)
        cell = ws[f'{column_letter}{row_th + 2}']
        cell.value = text


def render_all_card_shadows(document, ws):
    """全てのカードの影の描画
    """
    print('🔧　全てのカードの影の描画')

    # もし、柱のリストがあれば
    if 'pillars' in document and (pillars_list := document['pillars']):

        for pillar_dict in pillars_list:
            # もし、カードの辞書があれば
            if 'cards' in pillar_dict and (card_dict_list := pillar_dict['cards']):

                for card_dict in card_dict_list:
                    if 'shadowColor' in card_dict:
                        card_shadow_color = card_dict['shadowColor']

                        card_rect = Rectangle.from_dict(card_dict)

                        # 端子の影を描く
                        fill_rectangle(
                                ws=ws,
                                column_th=card_rect.left_obj.total_of_out_counts_th + OUT_COUNTS_THAT_CHANGE_INNING,
                                row_th=card_rect.top_obj.total_of_out_counts_th + OUT_COUNTS_THAT_CHANGE_INNING,
                                columns=card_rect.width_obj.total_of_out_counts_qty,
                                rows=card_rect.height_obj.total_of_out_counts_qty,
                                fill_obj=tone_and_color_name_to_fill_obj(card_shadow_color))


def render_all_cards(document, ws):
    """全てのカードの描画
    """
    print('🔧　全てのカードの描画')

    # もし、柱のリストがあれば
    if 'pillars' in document and (pillars_list := document['pillars']):

        for pillar_dict in pillars_list:

            # 柱と柱の隙間（隙間柱）は無視する
            if 'baseColor' not in pillar_dict or not pillar_dict['baseColor']:
                continue

            baseColor = pillar_dict['baseColor']
            card_list = pillar_dict['cards']

            for card_dict in card_list:

                card_rect = Rectangle.from_dict(card_dict)

                # ヘッダーの矩形の枠線を描きます
                draw_rectangle(
                        ws=ws,
                        column_th=card_rect.left_obj.total_of_out_counts_th,
                        row_th=card_rect.top_obj.total_of_out_counts_th,
                        columns=card_rect.width_obj.total_of_out_counts_qty,
                        rows=card_rect.height_obj.total_of_out_counts_qty)

                if 'paperStrips' in card_dict:
                    paper_strip_list = card_dict['paperStrips']

                    for index, paper_strip in enumerate(paper_strip_list):

                        # 短冊１行の描画
                        render_paper_strip(
                                ws=ws,
                                paper_strip=paper_strip,
                                column_th=card_rect.left_obj.total_of_out_counts_th,
                                row_th=index * OUT_COUNTS_THAT_CHANGE_INNING + card_rect.top_obj.total_of_out_counts_th,
                                columns=card_rect.width_obj.total_of_out_counts_qty,
                                rows=card_rect.height_obj.total_of_out_counts_qty)


def render_all_terminal_shadows(document, ws):
    """全ての端子の影の描画
    """
    print('🔧　全ての端子の影の描画')

    # もし、柱のリストがあれば
    if 'pillars' in document and (pillars_list := document['pillars']):

        for pillar_dict in pillars_list:
            # もし、端子のリストがあれば
            if 'terminals' in pillar_dict and (terminals_list := pillar_dict['terminals']):

                for terminal_dict in terminals_list:

                    terminal_rect = Rectangle.from_dict(terminal_dict)
                    terminal_shadow_color = terminal_dict['shadowColor']

                    # 端子の影を描く
                    fill_rectangle(
                            ws=ws,
                            column_th=terminal_rect.left_obj.total_of_out_counts_th + OUT_COUNTS_THAT_CHANGE_INNING,
                            row_th=terminal_rect.top_obj.total_of_out_counts_th + OUT_COUNTS_THAT_CHANGE_INNING,
                            columns=9,
                            rows=9,
                            fill_obj=tone_and_color_name_to_fill_obj(terminal_shadow_color))


def render_all_terminals(document, ws):
    """全ての端子の描画
    """
    print('🔧　全ての端子の描画')

    # もし、柱のリストがあれば
    if 'pillars' in document and (pillars_list := document['pillars']):

        for pillar_dict in pillars_list:
            # もし、端子のリストがあれば
            if 'terminals' in pillar_dict and (terminals_list := pillar_dict['terminals']):

                for terminal_dict in terminals_list:

                    terminal_pixel_art = terminal_dict['pixelArt']
                    terminal_rect = Rectangle.from_dict(terminal_dict)

                    if terminal_pixel_art == 'start':
                        # 始端のドット絵を描く
                        fill_start_terminal(
                            ws=ws,
                            column_th=terminal_rect.left_obj.total_of_out_counts_th,
                            row_th=terminal_rect.top_obj.total_of_out_counts_th)
                    
                    elif terminal_pixel_art == 'end':
                        # 終端のドット絵を描く
                        fill_end_terminal(
                            ws=ws,
                            column_th=terminal_rect.left_obj.total_of_out_counts_th,
                            row_th=terminal_rect.top_obj.total_of_out_counts_th)


def render_all_line_tape_shadows(document, ws):
    """全てのラインテープの影の描画
    """
    print('🔧　全てのラインテープの影の描画')

    # もし、ラインテープの配列があれば
    if 'lineTapes' in document and (line_tape_list := document['lineTapes']):

        for line_tape_dict in line_tape_list:
            for segment_dict in line_tape_dict['segments']:
                if 'shadowColor' in segment_dict and (line_tape_shadow_color := segment_dict['shadowColor']):
                    segment_rect = Rectangle.from_dict(segment_dict)

                    # 端子の影を描く
                    fill_rectangle(
                            ws=ws,
                            column_th=segment_rect.left_obj.total_of_out_counts_th + OUT_COUNTS_THAT_CHANGE_INNING,
                            row_th=segment_rect.top_obj.total_of_out_counts_th + OUT_COUNTS_THAT_CHANGE_INNING,
                            columns=segment_rect.width_obj.total_of_out_counts_qty,
                            rows=segment_rect.height_obj.total_of_out_counts_qty,
                            fill_obj=tone_and_color_name_to_fill_obj(line_tape_shadow_color))


def render_all_line_tapes(document, ws):
    """全てのラインテープの描画
    """
    print('🔧　全てのラインテープの描画')

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

                    segment_rect = Rectangle.from_dict(segment_dict)

                    # ラインテープを描く
                    fill_obj = tone_and_color_name_to_fill_obj(line_tape_color)
                    fill_rectangle(
                            ws=ws,
                            column_th=segment_rect.left_obj.total_of_out_counts_th,
                            row_th=segment_rect.top_obj.total_of_out_counts_th,
                            columns=segment_rect.width_obj.total_of_out_counts_qty,
                            rows=segment_rect.height_obj.total_of_out_counts_qty,
                            fill_obj=fill_obj)

                    # （あれば）アウトラインを描く
                    if line_tape_outline_color and line_tape_direction:
                        outline_fill_obj = tone_and_color_name_to_fill_obj(line_tape_outline_color)

                        # （共通処理）垂直方向
                        if line_tape_direction in ['from_here.falling_down', 'after_go_right.turn_falling_down', 'after_go_left.turn_up', 'after_go_left.turn_falling_down']:
                            # 左辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th - 1,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th + 1,
                                    columns=1,
                                    rows=segment_rect.height_obj.total_of_out_counts_qty - 2,
                                    fill_obj=outline_fill_obj)

                            # 右辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th + segment_rect.width_obj.total_of_out_counts_qty,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th + 1,
                                    columns=1,
                                    rows=segment_rect.height_obj.total_of_out_counts_qty - 2,
                                    fill_obj=outline_fill_obj)
                        
                        # （共通処理）水平方向
                        elif line_tape_direction in ['after_falling_down.turn_right', 'continue.go_right', 'after_falling_down.turn_left', 'continue.go_left', 'after_up.turn_right', 'from_here.go_right']:
                            # 上辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th + OUT_COUNTS_THAT_CHANGE_INNING,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th - 1,
                                    columns=segment_rect.width_obj.total_of_out_counts_qty - 2 * OUT_COUNTS_THAT_CHANGE_INNING,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                            # 下辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th + OUT_COUNTS_THAT_CHANGE_INNING,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th + segment_rect.height_obj.total_of_out_counts_qty,
                                    columns=segment_rect.width_obj.total_of_out_counts_qty - 2 * OUT_COUNTS_THAT_CHANGE_INNING,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                        # ここから落ちていく
                        if line_tape_direction == 'from_here.falling_down':
                            # 左辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th - 1,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th,
                                    columns=1,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                            # 右辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th + segment_rect.width_obj.total_of_out_counts_qty,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th,
                                    columns=1,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                        # 落ちたあと、右折
                        elif line_tape_direction == 'after_falling_down.turn_right':
                            # 左辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th - 1,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th - 1,
                                    columns=1,
                                    rows=2,
                                    fill_obj=outline_fill_obj)

                            # 下辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th - 1,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th + 1,
                                    columns=OUT_COUNTS_THAT_CHANGE_INNING + 1,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                        # そのまま右進
                        elif line_tape_direction == 'continue.go_right':
                            # 上辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th - OUT_COUNTS_THAT_CHANGE_INNING,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th - 1,
                                    columns=2 * OUT_COUNTS_THAT_CHANGE_INNING,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                            # 下辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th - OUT_COUNTS_THAT_CHANGE_INNING,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th + 1,
                                    columns=2 * OUT_COUNTS_THAT_CHANGE_INNING,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                        # 右進から落ちていく
                        elif line_tape_direction == 'after_go_right.turn_falling_down':
                            # 上辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th - OUT_COUNTS_THAT_CHANGE_INNING,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th - 1,
                                    columns=2 * OUT_COUNTS_THAT_CHANGE_INNING,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                            # 下辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th - OUT_COUNTS_THAT_CHANGE_INNING,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th + 1,
                                    columns=OUT_COUNTS_THAT_CHANGE_INNING,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                            # 右辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th + segment_rect.width_obj.total_of_out_counts_qty,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th - 1,
                                    columns=1,
                                    rows=2,
                                    fill_obj=outline_fill_obj)

                        # 落ちたあと左折
                        elif line_tape_direction == 'after_falling_down.turn_left':
                            # 右辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th + segment_rect.width_obj.total_of_out_counts_qty,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th - 1,
                                    columns=1,
                                    rows=2,
                                    fill_obj=outline_fill_obj)

                            # 下辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th + segment_rect.width_obj.total_of_out_counts_qty - OUT_COUNTS_THAT_CHANGE_INNING,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th + 1,
                                    columns=OUT_COUNTS_THAT_CHANGE_INNING + 1,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                        # そのまま左進
                        elif line_tape_direction == 'continue.go_left':
                            # 上辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th + OUT_COUNTS_THAT_CHANGE_INNING,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th - 1,
                                    columns=segment_rect.width_obj.total_of_out_counts_qty,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                            # 下辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th + OUT_COUNTS_THAT_CHANGE_INNING,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th + 1,
                                    columns=segment_rect.width_obj.total_of_out_counts_qty,
                                    rows=1,
                                    fill_obj=outline_fill_obj)
                        
                        # 左進から上っていく
                        elif line_tape_direction == 'after_go_left.turn_up':
                            # 下辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th + segment_rect.height_obj.total_of_out_counts_qty,
                                    columns=2 * OUT_COUNTS_THAT_CHANGE_INNING,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                            # 左辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th - 1,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th + segment_rect.height_obj.total_of_out_counts_qty - 2,
                                    columns=1,
                                    rows=3,
                                    fill_obj=outline_fill_obj)
                            
                            # 右辺（横長）を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th + OUT_COUNTS_THAT_CHANGE_INNING,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th + segment_rect.height_obj.total_of_out_counts_qty - 2,
                                    columns=OUT_COUNTS_THAT_CHANGE_INNING,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                        # 上がってきて右折
                        elif line_tape_direction == 'after_up.turn_right':
                            # 左辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th - 1,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th,
                                    columns=1,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                            # 上辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th - 1,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th - 1,
                                    columns=OUT_COUNTS_THAT_CHANGE_INNING + 1,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                        # 左進から落ちていく
                        elif line_tape_direction == 'after_go_left.turn_falling_down':
                            # 上辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th - 1,
                                    columns=2 * OUT_COUNTS_THAT_CHANGE_INNING,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                            # 左辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th - 1,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th - 1,
                                    columns=1,
                                    rows=segment_rect.height_obj.total_of_out_counts_qty,
                                    fill_obj=outline_fill_obj)

                            # 右辺（横長）を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th + OUT_COUNTS_THAT_CHANGE_INNING + 1,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th + 1,
                                    columns=OUT_COUNTS_THAT_CHANGE_INNING - 1,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                        # ここから右進
                        elif line_tape_direction == 'from_here.go_right':
                            # 上辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th - 1,
                                    columns=OUT_COUNTS_THAT_CHANGE_INNING,
                                    rows=1,
                                    fill_obj=outline_fill_obj)

                            # 下辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th + 1,
                                    columns=OUT_COUNTS_THAT_CHANGE_INNING,
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

                pillar_rect = Rectangle.from_dict(pillar_dict)
                base_color = pillar_dict['baseColor']

                # もし、矩形の中に、指定の点が含まれたなら
                if pillar_rect.left_obj.total_of_out_counts_th <= column_th and column_th < pillar_rect.left_obj.total_of_out_counts_th + pillar_rect.width_obj.total_of_out_counts_qty and \
                    pillar_rect.top_obj.total_of_out_counts_th <= row_th and row_th < pillar_rect.top_obj.total_of_out_counts_th + pillar_rect.height_obj.total_of_out_counts_qty:

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
                            card_rect = Rectangle.from_dict(card_dict)

                            # 影に自動が設定されていたら、解決する
                            if solved_tone_and_color_name := resolve_auto_shadow(
                                    document=document,
                                    column_th=card_rect.left_obj.total_of_out_counts_th + OUT_COUNTS_THAT_CHANGE_INNING,
                                    row_th=card_rect.top_obj.total_of_out_counts_th + OUT_COUNTS_THAT_CHANGE_INNING):
                                card_dict['shadowColor'] = solved_tone_and_color_name

            # もし、端子のリストがあれば
            if 'terminals' in pillar_dict and (terminals_list := pillar_dict['terminals']):

                for terminal_dict in terminals_list:
                    if 'shadowColor' in terminal_dict and (terminal_shadow_color := terminal_dict['shadowColor']):

                        if terminal_shadow_color == 'auto':
                            terminal_rect = Rectangle.from_dict(terminal_dict)

                            # 影に自動が設定されていたら、解決する
                            if solved_tone_and_color_name := resolve_auto_shadow(
                                    document=document,
                                    column_th=terminal_rect.left_obj.total_of_out_counts_th + OUT_COUNTS_THAT_CHANGE_INNING,
                                    row_th=terminal_rect.top_obj.total_of_out_counts_th + OUT_COUNTS_THAT_CHANGE_INNING):
                                terminal_dict['shadowColor'] = solved_tone_and_color_name

    # もし、ラインテープのリストがあれば
    if 'lineTapes' in document and (line_tape_list := document['lineTapes']):

        for line_tape_dict in line_tape_list:
            # もし、セグメントのリストがあれば
            if 'segments' in line_tape_dict and (segment_list := line_tape_dict['segments']):

                for segment_dict in segment_list:
                    if 'shadowColor' in segment_dict and (segment_shadow_color := segment_dict['shadowColor']) and segment_shadow_color == 'auto':
                        segment_rect = Rectangle.from_dict(segment_dict)

                        # NOTE 影が指定されているということは、浮いているということでもある

                        # 影に自動が設定されていたら、解決する
                        if solved_tone_and_color_name := resolve_auto_shadow(
                                document=document,
                                column_th=segment_rect.left_obj.total_of_out_counts_th + OUT_COUNTS_THAT_CHANGE_INNING,
                                row_th=segment_rect.top_obj.total_of_out_counts_th + OUT_COUNTS_THAT_CHANGE_INNING):
                            segment_dict['shadowColor'] = solved_tone_and_color_name


def split_segment_by_pillar(document, line_tape_segment_list, line_tape_segment_dict):
    """柱を跨ぐとき、ラインテープを分割します
    NOTE 柱は左から並んでいるものとする
    NOTE 柱の縦幅は十分に広いものとする
    NOTE テープは浮いています
    """

    new_segment_list = []

    #print('🔧　柱を跨ぐとき、ラインテープを分割します')
    segment_rect = Rectangle.from_dict(line_tape_segment_dict)

    direction = line_tape_segment_dict['direction']

    splitting_segments = []


    # 右進でも、左進でも、同じコードでいけるようだ
    if direction in ['after_falling_down.turn_right', 'after_up.turn_right', 'from_here.go_right', 'after_falling_down.turn_left']:

        # もし、柱のリストがあれば
        if 'pillars' in document and (pillars_list := document['pillars']):

            # 各柱
            for pillar_dict in pillars_list:
                pillar_rect = Rectangle.from_dict(pillar_dict)

                # とりあえず、ラインテープの左端と右端の内側に、柱の右端があるか判定
                if segment_rect.left_obj.total_of_out_counts_th < pillar_rect.right_obj.total_of_out_counts_th and pillar_rect.right_obj.total_of_out_counts_th < segment_rect.right_obj.total_of_out_counts_th:
                    # 既存のセグメントを削除
                    line_tape_segment_list.remove(line_tape_segment_dict)

                    # 左側のセグメントを新規作成し、新リストに追加
                    # （計算を簡単にするため）width は使わず right を使う
                    left_segment_dict = dict(line_tape_segment_dict)
                    left_segment_dict.pop('width', None)
                    left_segment_dict['right'] = InningsPitched(pillar_rect.right_obj.var_value).offset(-1).var_value
                    new_segment_list.append(left_segment_dict)

                    # 右側のセグメントを新規作成し、既存リストに追加
                    # （計算を簡単にするため）width は使わず right を使う
                    right_segment_dict = dict(line_tape_segment_dict)
                    right_segment_dict.pop('width', None)
                    right_segment_dict['left'] = pillar_rect.right_obj.offset(-1).var_value
                    right_segment_dict['right'] = segment_rect.right_obj.var_value
                    line_tape_segment_list.append(right_segment_dict)
                    line_tape_segment_dict = right_segment_dict          # 入れ替え

    
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
