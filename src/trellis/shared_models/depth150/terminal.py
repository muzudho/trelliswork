from ..depth130 import Rectangle
from ..depth140 import Canvas


class Terminal():
    """端子
    """


    def from_dict(terminal_dict):

        bounds_obj = None
        if 'bounds' in terminal_dict and (bounds_dict := terminal_dict['bounds']):
            bounds_obj = Rectangle.from_dict(bounds_dict)

        return Canvas(
                bounds_obj=bounds_obj)


    def __init__(self, bounds_obj):
        self._bounds_obj = bounds_obj


    @property
    def bounds_obj(self):
        return self._bounds_obj
