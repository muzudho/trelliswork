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
        return InningsPitched(integer_part=integer_part, decimal_part=decimal_part)


    @staticmethod
    def from_var_value(var_value):

        try:
            # "100" が来たら 100 にする
            var_value = int(var_value)
        except ValueError:
            pass

        if isinstance(var_value, int):
            return InningsPitched(
                    integer_part=var_value,
                    decimal_part=0)

        elif isinstance(var_value, str):
            integer_part, decimal_part = map(int, var_value.split('o', 2))
            return InningsPitched(
                    integer_part=integer_part,
                    decimal_part=decimal_part)

        else:
            raise ValueError(f'{type(var_value)=} {var_value=}')

        return InningsPitched(var_value)


    def __init__(self, integer_part, decimal_part):
        self._integer_part = integer_part
        self._decimal_part = decimal_part

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
        r = InningsPitched.from_var_value(var_value)  # Right operand
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
        self._bottom_obj = None


    def _calculate_right(self):
        sum_decimal_part = self._left_obj.decimal_part + self._width_obj.decimal_part
        self._right_obj = InningsPitched.from_integer_and_decimal_part(
                integer_part=self._left_obj.integer_part + self._width_obj.integer_part + sum_decimal_part // OUT_COUNTS_THAT_CHANGE_INNING,
                decimal_part=sum_decimal_part % OUT_COUNTS_THAT_CHANGE_INNING)


    def _calculate_bottom(self):
        sum_decimal_part = self._top_obj.decimal_part + self._height_obj.decimal_part
        self._bottom_obj = InningsPitched.from_integer_and_decimal_part(
                integer_part=self._top_obj.integer_part + self._height_obj.integer_part + sum_decimal_part // OUT_COUNTS_THAT_CHANGE_INNING,
                decimal_part=sum_decimal_part % OUT_COUNTS_THAT_CHANGE_INNING)


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
    def bottom_obj(self):
        """矩形の下位置
        """
        if not self._bottom_obj:
            self._calculate_bottom()
        return self._bottom_obj


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
