class Point():
    """位置
    """


    @staticmethod
    def from_dict(point_dict):
        """辞書を元に生成
        """

        try:
            first_x = point_dict['x']
        except:
            print(f'ERROR: Rectangle.from_dict: {point_dict=}')
            raise

        second_x = 0
        if isinstance(first_x, str):
            first_x, second_x = map(int, first_x.split('o', 2))

        first_y = point_dict['y']
        second_y = 0
        if isinstance(first_y, str):
            first_y, second_y = map(int, first_y.split('o', 2))

        return Point(
                first_x=first_x,
                second_x=second_x,
                first_y=first_y,
                second_y=second_y)


    def __init__(self, first_x, second_x, first_y, second_y):
        """初期化
        """
        self._x_obj = InningsPitched.from_integer_and_decimal_part(integer_part=first_x, decimal_part=second_x)
        self._y_obj = InningsPitched.from_integer_and_decimal_part(integer_part=first_y, decimal_part=second_y)


    @property
    def x_obj(self):
        return self._x_obj


    @property
    def y_obj(self):
        return self._y_obj
