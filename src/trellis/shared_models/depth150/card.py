from ..depth130 import Rectangle
from ..depth140 import Canvas


class Card():
    """カード
    """


    def from_dict(card_dict):

        bounds_obj = None
        if 'varBounds' in card_dict and (o1_bounds_dict := card_dict['varBounds']):
            bounds_obj = Rectangle.from_var_bounds_dict(o1_bounds_dict)
        elif 'bounds' in card_dict and (o2_bounds_dict := card_dict['bounds']):
            bounds_obj = Rectangle.from_bounds_dict(o2_bounds_dict)

        return Canvas(
                bounds_obj=bounds_obj)


    def __init__(self, bounds_obj):
        self._bounds_obj = bounds_obj


    @property
    def bounds_obj(self):
        return self._bounds_obj
