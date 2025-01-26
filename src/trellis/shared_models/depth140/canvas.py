from ..depth130 import Rectangle


class Canvas():
    """キャンバス
    """


    def from_dict(canvas_dict):

        bounds_obj = None
        if 'varBounds' in canvas_dict and (bounds_dict := canvas_dict['varBounds']):
            bounds_obj = Rectangle.from_var_bounds_dict(bounds_dict)

        return Canvas(
                bounds_obj=bounds_obj)


    def __init__(self, bounds_obj):
        self._bounds_obj = bounds_obj


    @property
    def bounds_obj(self):
        return self._bounds_obj
