from ..depth130 import VarRectangle
from ..depth140 import Canvas


class Terminal():
    """端子
    """


    def from_dict(terminal_dict):

        bounds_obj = None
        if 'varBounds' in terminal_dict and (o1_bounds_dict := terminal_dict['varBounds']):
            bounds_obj = VarRectangle.from_var_bounds_dict(o1_bounds_dict)
        elif 'bounds' in terminal_dict and (o2_bounds_dict := terminal_dict['bounds']):
            bounds_obj = VarRectangle.from_bounds_dict(o2_bounds_dict)

        return Canvas(
                bounds_obj=bounds_obj)


    def __init__(self, bounds_obj):
        self._bounds_obj = bounds_obj


    @property
    def bounds_obj(self):
        return self._bounds_obj
