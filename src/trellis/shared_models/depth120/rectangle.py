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
            left = bounds_dict['left']
        except:
            print(f'ERROR: VarRectangle.from_bounds_dict: {bounds_dict=}')
            raise

        top = bounds_dict['top']

        # right は、その数を含まない。
        # right が指定されていれば、 width より優先する
        if 'right' in bounds_dict:
            right = bounds_dict['right']
            width = right - left

        else:
            width = bounds_dict['width']

        # bottom は、その数を含まない。
        # bottom が指定されていれば、 width より優先する
        if 'bottom' in bounds_dict:
            bottom = bounds_dict['bottom']
            height = bottom - top

        else:
            height = bounds_dict['height']

        return VarRectangle(
                left=left,
                top=top,
                width=width,
                height=height)


    def __init__(self, left, top, width, height):
        """初期化
        """
        self._left = left
        self._top = top
        self._width = width
        self._height = height
        self._right = None
        self._bottom = None


    def _calculate_right(self):
        self._right = self._left + self._width


    def _calculate_bottom(self):
        self._bottom = self._top + self._height


    @property
    def left(self):
        return self._left


    @property
    def right(self):
        """矩形の右位置
        """
        if not self._right:
            self._calculate_right()
        return self._right


    @property
    def top(self):
        return self._top


    @property
    def bottom(self):
        """矩形の下位置
        """
        if not self._bottom:
            self._calculate_bottom()
        return self._bottom


    @property
    def width(self):
        return self._width


    @property
    def height(self):
        return self._height


    def to_ltwh_dict(self):
        """left, top, width, height を含む辞書を作成します
        """

        left = self._left
        if isinstance(left, str):
            left = f'"{left}"'

        top = self._top
        if isinstance(top, str):
            top = f'"{top}"'

        width = self._width
        if isinstance(width, str):
            width = f'"{width}"'

        height = self._height
        if isinstance(height, str):
            height = f'"{height}"'

        return {
            "left": left,
            "top": top,
            "width": width,
            "height": height
        }


    def to_lrtb_dict(self):
        """left, right, top, bottom を含む辞書を作成します
        """

        left = self._left
        if isinstance(left, str):
            left = f'"{left}"'

        right = self._right
        if isinstance(right, str):
            right = f'"{right}"'

        top = self._top
        if isinstance(top, str):
            top = f'"{top}"'

        bottom = self._bottom
        if isinstance(bottom, str):
            bottom = f'"{bottom}"'

        return {
            "left": left,
            "right": right,
            "top": top,
            "bottom": bottom
        }
