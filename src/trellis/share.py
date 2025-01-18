import os
import openpyxl as xl
from openpyxl.styles import PatternFill, Font
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
    def from_dict(rect_dict):
        """ラインテープのセグメントの矩形情報を取得
        """

        try:
            main_left = rect_dict['left']
        except:
            print(f'ERROR: Rectangle.from_dict: {rect_dict=}')
            raise

        sub_left = 0
        if isinstance(main_left, str):
            main_left, sub_left = map(int, main_left.split('o', 2))

        main_top = rect_dict['top']
        sub_top = 0
        if isinstance(main_top, str):
            main_top, sub_top = map(int, main_top.split('o', 2))

        # right は、その数を含まない。
        # right が指定されていれば、 width より優先する
        if 'right' in rect_dict:
            right = rect_dict['right']
            sub_right = 0
            if isinstance(right, str):
                right, sub_right = map(int, right.split('o', 2))

            main_width = right - main_left
            sub_width = sub_right - sub_left

        else:
            main_width = rect_dict['width']
            sub_width = 0
            if isinstance(main_width, str):
                main_width, sub_width = map(int, main_width.split('o', 2))

        # bottom は、その数を含まない。
        # bottom が指定されていれば、 width より優先する
        if 'bottom' in rect_dict:
            bottom = rect_dict['bottom']
            sub_bottom = 0
            if isinstance(bottom, str):
                bottom, sub_bottom = map(int, bottom.split('o', 2))

            main_height = bottom - main_top
            sub_height = sub_bottom - sub_top

        else:
            main_height = rect_dict['height']
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
none_pattern_fill = PatternFill(patternType=None)
# エクセルの色システム（勝手に作ったったもの）
web_safe_color_code_dict = {
    'xl_theme' : {
        'xl_white' : '#FFFFFF',
        'xl_black' : '#000000',
        'xl_red_gray' : '#E7E6E6',
        'xl_blue_gray' : '#44546A',
        'xl_blue' : '#5B9BD5',
        'xl_red' : '#ED7D31',
        'xl_gray' : '#A5A5A5',
        'xl_yellow' : '#FFC000',
        'xl_naviy' : '#4472C4',
        'xl_green' : '#70AD47',
    },
    'xl_pale' : {
        'xl_white' : '#F2F2F2',
        'xl_black' : '#808080',
        'xl_red_gray' : '#AEAAAA',
        'xl_blue_gray' : '#D6DCE4',
        'xl_blue' : '#DDEBF7',
        'xl_red' : '#FCE4D6',
        'xl_gray' : '#EDEDED',
        'xl_yellow' : '#FFF2CC',
        'xl_naviy' : '#D9E1F2',
        'xl_green' : '#E2EFDA',
    },
    'xl_light' : {
        'xl_white' : '#D9D9D9',
        'xl_black' : '#595959',
        'xl_red_gray' : '#757171',
        'xl_blue_gray' : '#ACB9CA',
        'xl_blue' : '#BDD7EE',
        'xl_red' : '#F8CBAD',
        'xl_gray' : '#DBDBDB',
        'xl_yellow' : '#FFE699',
        'xl_naviy' : '#B4C6E7',
        'xl_green' : '#C6E0B4',
    },
    'xl_soft' : {
        'xl_white' : '#BFBFBF',
        'xl_black' : '#404040',
        'xl_red_gray' : '#3A3838',
        'xl_blue_gray' : '#8497B0',
        'xl_blue' : '#9BC2E6',
        'xl_red' : '#F4B084',
        'xl_gray' : '#C9C9C9',
        'xl_yellow' : '#FFD966',
        'xl_naviy' : '#8EA9DB',
        'xl_green' : '#A9D08E',
    },
    'xl_strong' : {
        'xl_white' : '#A6A6A6',
        'xl_black' : '#262626',
        'xl_red_gray' : '#3A3838',
        'xl_blue_gray' : '#333F4F',
        'xl_blue' : '#2F75B5',
        'xl_red' : '#C65911',
        'xl_gray' : '#7B7B7B',
        'xl_yellow' : '#BF8F00',
        'xl_naviy' : '#305496',
        'xl_green' : '#548235',
    },
    'xl_deep' : {
        'xl_white' : '#808080',
        'xl_black' : '#0D0D0D',
        'xl_red_gray' : '#161616',
        'xl_blue_gray' : '#161616',
        'xl_blue' : '#1F4E78',
        'xl_red' : '#833C0C',
        'xl_gray' : '#525252',
        'xl_yellow' : '#806000',
        'xl_naviy' : '#203764',
        'xl_green' : '#375623',
    },
    'xl_standard' : {
        'xl_red' : '#C00000',
        'xl_red' : '#FF0000',
        'xl_orange' : '#FFC000',
        'xl_yellow' : '#FFFF00',
        'xl_yellow_green' : '#92D050',
        'xl_green' : '#00B050',
        'xl_dodger_blue' : '#00B0F0',
        'xl_blue' : '#0070C0',
        'xl_naviy' : '#002060',
        'xl_violet' : '#7030A0',
    }
}


def web_safe_color_code_to_xl(web_safe_color_code):
    """頭の `#` を外します
    """
    return web_safe_color_code[1:]


def tone_and_color_name_to_web_safe_color_code(tone_and_color_name):
    """トーン名・色名をウェブ・セーフ・カラーの１６進文字列の色コードに変換します
    """

    # 色が指定されていないとき、この関数を呼び出してはいけません
    if tone_and_color_name is None:
        raise Exception(f'tone_and_color_name_to_web_safe_color_code: 色が指定されていません')

    # 背景色を［なし］にします。透明（transparent）で上書きするのと同じです
    if tone_and_color_name == 'paper_color':
        raise Exception(f'tone_and_color_name_to_web_safe_color_code: 透明色には対応していません')

    # ［auto］は自動で影の色を設定する機能ですが、その機能をオフにしているときは、とりあえず黒色にします
    if tone_and_color_name == 'auto':
        return web_safe_color_code_dict['xl_theme']['xl_black']

    # `#` で始まるなら、ウェブセーフカラーとして扱う
    if tone_and_color_name.startswith('#'):
        return tone_and_color_name


    try:
        tone, color = tone_and_color_name.split('.', 2)
    except:
        print(f'tone_and_color_name_to_web_safe_color_code: tone.color の形式でない {tone_and_color_name=}')
        raise


    tone = tone.strip()
    color = color.strip()

    if tone in web_safe_color_code_dict:
        if color in web_safe_color_code_dict[tone]:
            return web_safe_color_code_dict[tone][color]

    print(f'tone_and_color_name_to_web_safe_color_code: 色がない {tone_and_color_name=}')
    return None


def tone_and_color_name_to_fill_obj(tone_and_color_name):
    """トーン名・色名を FillPattern オブジェクトに変換します
    """

    # 色が指定されていないとき、この関数を呼び出してはいけません
    if tone_and_color_name is None:
        raise Exception(f'tone_and_color_name_to_fill_obj: 色が指定されていません')

    # 背景色を［なし］にします。透明（transparent）で上書きするのと同じです
    if tone_and_color_name == 'paper_color':
        return none_pattern_fill

    # ［auto］は自動で影の色を設定する機能ですが、その機能をオフにしているときは、とりあえず黒色にします
    if tone_and_color_name == 'auto':
        return PatternFill(
                patternType='solid',
                fgColor=web_safe_color_code_to_xl(web_safe_color_code_dict['xl_theme']['xl_black']))

    try:
        tone, color = tone_and_color_name.split('.', 2)
    except:
        print(f'ERROR: {tone_and_color_name=}')
        raise

    tone = tone.strip()
    color = color.strip()

    if tone in web_safe_color_code_dict:
        if color in web_safe_color_code_dict[tone]:
            return PatternFill(
                    patternType='solid',
                    fgColor=web_safe_color_code_to_xl(web_safe_color_code_dict[tone][color]))

    print(f'tone_and_color_name_to_fill_obj: 色がない {tone_and_color_name=}')
    return none_pattern_fill


###################
# MARK: XlAlignment
###################
class XlAlignment():
    """Excel 用テキストの位置揃え
    """


    @staticmethod
    def from_dict(xl_alignment_dict):
        """辞書を元に生成

        📖 [openpyxl.styles.alignment module](https://openpyxl.readthedocs.io/en/latest/api/openpyxl.styles.alignment.html)
        horizontal: Value must be one of {‘fill’, ‘left’, ‘distributed’, ‘justify’, ‘center’, ‘general’, ‘centerContinuous’, ‘right’}
        vertical: Value must be one of {‘distributed’, ‘justify’, ‘center’, ‘bottom’, ‘top’}
        """
        xl_horizontal = None
        xl_vertical = None
        if 'xl_horizontal' in xl_alignment_dict:
            xl_horizontal = xl_alignment_dict['xl_horizontal']

        if 'xl_vertical' in xl_alignment_dict:
            xl_vertical = xl_alignment_dict['xl_vertical']

        return XlAlignment(
                xl_horizontal=xl_horizontal,
                xl_vertical=xl_vertical)


    def __init__(self, xl_horizontal, xl_vertical):
        self._xl_horizontal = xl_horizontal
        self._xl_vertical = xl_vertical


    @property
    def xl_horizontal(self):
        return self._xl_horizontal


    @property
    def xl_vertical(self):
        return self._xl_vertical


##############
# MARK: XlFont
##############
class XlFont():
    """Excel 用フォント
    """


    @staticmethod
    def from_dict(xl_font_dict):
        """辞書を元に生成
        """
        web_safe_color_code = None
        if 'color' in xl_font_dict:
            web_safe_color_code = tone_and_color_name_to_web_safe_color_code(xl_font_dict['color'])

        return XlFont(
                web_safe_color_code=web_safe_color_code)


    def __init__(self, web_safe_color_code):
        self._web_safe_color_code = web_safe_color_code


    @property
    def web_safe_color_code(self):
        return self._web_safe_color_code


    @property
    def color_code_for_xl(self):
        return web_safe_color_code_to_xl(self._web_safe_color_code)


##############
# MARK: Canvas
##############
class Canvas():
    """キャンバス
    """


    def from_dict(canvas_dict):

        rect_obj = None
        if 'rect' in canvas_dict and (rect_dict := canvas_dict['rect']):
            rect_obj = Rectangle.from_dict(rect_dict)

        return Canvas(
                rect_obj=rect_obj)


    def __init__(self, rect_obj):
        self._rect_obj = rect_obj


    @property
    def rect_obj(self):
        return self._rect_obj


##############
# MARK: Pillar
##############
class Pillar():
    """柱
    """


    def from_dict(pillar_dict):

        rect_obj = None
        if 'rect' in pillar_dict and (rect_dict := pillar_dict['rect']):
            rect_obj = Rectangle.from_dict(rect_dict)

        # FIXME: if 'baseColor' in pillar_dict and (tone_and_color_name := pillar_dict['baseColor']):


        return Canvas(
                rect_obj=rect_obj)


    def __init__(self, rect_obj):
        self._rect_obj = rect_obj


    @property
    def rect_obj(self):
        return self._rect_obj


############
# MARK: Card
############
class Card():
    """カード
    """


    def from_dict(card_dict):

        rect_obj = None
        if 'rect' in card_dict and (rect_dict := card_dict['rect']):
            rect_obj = Rectangle.from_dict(rect_dict)

        # FIXME: if 'baseColor' in pillar_dict and (tone_and_color_name := pillar_dict['baseColor']):


        return Canvas(
                rect_obj=rect_obj)


    def __init__(self, rect_obj):
        self._rect_obj = rect_obj


    @property
    def rect_obj(self):
        return self._rect_obj
