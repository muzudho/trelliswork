from ..depth110 import Share
from ..depth120 import InningsPitched


class VarRectangle():
    """矩形
    """


    @staticmethod
    def from_bounds_dict(bounds_dict):
        """矩形情報を取得
        left,top,right,bottom,width,height の単位はそれぞれアウトカウント。
        """

        try:
            sub_left = bounds_dict['left']
        except:
            print(f'ERROR: VarRectangle.from_bounds_dict: {bounds_dict=}')
            raise

        sub_top = bounds_dict['top']

        # right は、その数を含まない。
        # right が指定されていれば、 width より優先する
        if 'right' in bounds_dict:
            sub_right = bounds_dict['right']
            sub_width = sub_right - sub_left

        else:
            sub_width = bounds_dict['width']

        # bottom は、その数を含まない。
        # bottom が指定されていれば、 width より優先する
        if 'bottom' in bounds_dict:
            sub_bottom = bounds_dict['bottom']
            sub_height = sub_bottom - sub_top

        else:
            sub_height = bounds_dict['height']

        return VarRectangle(
                main_left=0,
                sub_left=sub_left,
                main_top=0,
                sub_top=sub_top,
                main_width=0,
                sub_width=sub_width,
                main_height=0,
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
                integer_part=self._left_obj.integer_part + self._width_obj.integer_part + sum_decimal_part // Share.OUT_COUNTS_THAT_CHANGE_INNING,
                decimal_part=sum_decimal_part % Share.OUT_COUNTS_THAT_CHANGE_INNING)


    def _calculate_bottom(self):
        sum_decimal_part = self._top_obj.decimal_part + self._height_obj.decimal_part
        self._bottom_obj = InningsPitched.from_integer_and_decimal_part(
                integer_part=self._top_obj.integer_part + self._height_obj.integer_part + sum_decimal_part // Share.OUT_COUNTS_THAT_CHANGE_INNING,
                decimal_part=sum_decimal_part % Share.OUT_COUNTS_THAT_CHANGE_INNING)


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


    def to_var_ltwh_dict(self):
        """left, top, width, height を含む辞書を作成します
        """

        left = self._left_obj.var_value
        if isinstance(left, str):
            left = f'"{left}"'

        top = self._top_obj.var_value
        if isinstance(top, str):
            top = f'"{top}"'

        width = self._width_obj.var_value
        if isinstance(width, str):
            width = f'"{width}"'

        height = self._height_obj.var_value
        if isinstance(height, str):
            height = f'"{height}"'

        return {
            "left": left,
            "top": top,
            "width": width,
            "height": height
        }


    def to_var_lrtb_dict(self):
        """left, right, top, bottom を含む辞書を作成します
        """

        left = self._left_obj.var_value
        if isinstance(left, str):
            left = f'"{left}"'

        right = self._right_obj.var_value
        if isinstance(right, str):
            right = f'"{right}"'

        top = self._top_obj.var_value
        if isinstance(top, str):
            top = f'"{top}"'

        bottom = self._bottom_obj.var_value
        if isinstance(bottom, str):
            bottom = f'"{bottom}"'

        return {
            "left": left,
            "right": right,
            "top": top,
            "bottom": bottom
        }
