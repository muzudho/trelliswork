from .innings_pitched import InningsPitched
from .share import Share


class Rectangle():
    """矩形
    """


    @staticmethod
    def from_dict(bounds_dict):
        """ラインテープのセグメントの矩形情報を取得
        """

        try:
            main_left = bounds_dict['left']
        except:
            print(f'ERROR: Rectangle.from_dict: {bounds_dict=}')
            raise

        sub_left = 0
        if isinstance(main_left, str):
            main_left, sub_left = map(int, main_left.split('o', 2))

        main_top = bounds_dict['top']
        sub_top = 0
        if isinstance(main_top, str):
            main_top, sub_top = map(int, main_top.split('o', 2))

        # right は、その数を含まない。
        # right が指定されていれば、 width より優先する
        if 'right' in bounds_dict:
            right = bounds_dict['right']
            sub_right = 0
            if isinstance(right, str):
                right, sub_right = map(int, right.split('o', 2))

            main_width = right - main_left
            sub_width = sub_right - sub_left

        else:
            main_width = bounds_dict['width']
            sub_width = 0
            if isinstance(main_width, str):
                main_width, sub_width = map(int, main_width.split('o', 2))

        # bottom は、その数を含まない。
        # bottom が指定されていれば、 width より優先する
        if 'bottom' in bounds_dict:
            bottom = bounds_dict['bottom']
            sub_bottom = 0
            if isinstance(bottom, str):
                bottom, sub_bottom = map(int, bottom.split('o', 2))

            main_height = bottom - main_top
            sub_height = sub_bottom - sub_top

        else:
            main_height = bounds_dict['height']
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
