from ..depth130 import Rectangle
from ..depth140 import Canvas


class Card():
    """カード
    """


    def from_dict(card_dict):

        bounds_obj = None
        if 'varBounds' in card_dict and (bounds_dict := card_dict['varBounds']):
            bounds_obj = Rectangle.from_var_bounds_dict(bounds_dict)

        return Canvas(
                bounds_obj=bounds_obj)


    def __init__(self, bounds_obj):
        self._bounds_obj = bounds_obj


    @property
    def bounds_obj(self):
        return self._bounds_obj
