from ..depth110 import Share


class Rectangle():
    """矩形
    """


    @staticmethod
    def from_bounds_dict(bounds_dict):
        """矩形情報を取得
        left,top,right,bottom,width,height の単位はそれぞれアウトカウント。
        """

        try:
            left_qty = bounds_dict['left']
        except:
            print(f'ERROR: Rectangle.from_bounds_dict: {bounds_dict=}')
            raise

        top_qty = bounds_dict['top']

        # right は、その数を含まない。
        # right が指定されていれば、 width より優先する
        if 'right' in bounds_dict:
            right_qty = bounds_dict['right']
            width = right_qty - left_qty

        else:
            width = bounds_dict['width']

        # bottom は、その数を含まない。
        # bottom が指定されていれば、 width より優先する
        if 'bottom' in bounds_dict:
            bottom_qty = bounds_dict['bottom']
            height = bottom_qty - top_qty

        else:
            height = bounds_dict['height']

        return Rectangle(
                left_qty=left_qty,
                top_qty=top_qty,
                width=width,
                height=height)


    def __init__(self, left_qty, top_qty, width, height):
        """初期化
        """
        self._left_qty = left_qty
        self._top_qty = top_qty
        self._width = width
        self._height = height
        self._right_qty = None
        self._bottom_qty = None


    def _calculate_right(self):
        self._right_qty = self._left_qty + self._width


    def _calculate_bottom(self):
        self._bottom_qty = self._top_qty + self._height


    @property
    def left_qty(self):
        return self._left_qty


    @property
    def left_th(self):
        """Excel の列番号が 1 相当から始まるので、それに合わせるのに使う"""
        return self._left_qty + 1


    @property
    def right_qty(self):
        """矩形の右位置
        """
        if not self._right_qty:
            self._calculate_right()
        return self._right_qty


    @property
    def right_th(self):
        """矩形の右位置。Excel の列番号が 1 相当から始まるので、それに合わせるのに使う"""
        return self.right_qty + 1


    @property
    def top_qty(self):
        return self._top_qty


    @property
    def top_th(self):
        """Excel の行番号が 1 から始まるので、それに合わせるのに使う"""
        return self._top_qty + 1


    @property
    def bottom_qty(self):
        """矩形の下位置
        """
        if not self._bottom_qty:
            self._calculate_bottom()
        return self._bottom_qty


    @property
    def bottom_th(self):
        """矩形の下位置。Excel の行番号が 1 から始まるので、それに合わせるのに使う
        """
        return self.bottom_qty + 1


    @property
    def width(self):
        return self._width


    @property
    def height(self):
        return self._height


    def to_ltwh_dict(self):
        """left, top, width, height を含む辞書を作成します
        """

        left_qty = self._left_qty
        if isinstance(left_qty, str):
            left_qty = f'"{left_qty}"'

        top_qty = self._top_qty
        if isinstance(top_qty, str):
            top_qty = f'"{top_qty}"'

        width = self._width
        if isinstance(width, str):
            width = f'"{width}"'

        height = self._height
        if isinstance(height, str):
            height = f'"{height}"'

        return {
            "left": left_qty,
            "top": top_qty,
            "width": width,
            "height": height
        }


    def to_lrtb_dict(self):
        """left, right, top, bottom を含む辞書を作成します
        """

        left_qty = self._left_qty
        if isinstance(left_qty, str):
            left_qty = f'"{left_qty}"'

        right_qty = self._right_qty
        if isinstance(right_qty, str):
            right_qty = f'"{right_qty}"'

        top_qty = self._top_qty
        if isinstance(top_qty, str):
            top_qty = f'"{top_qty}"'

        bottom_qty = self._bottom_qty
        if isinstance(bottom_qty, str):
            bottom_qty = f'"{bottom_qty}"'

        return {
            "left": left_qty,
            "right": right_qty,
            "top": top_qty,
            "bottom": bottom_qty
        }
