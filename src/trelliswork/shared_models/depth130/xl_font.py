from ..depth110 import WebSafeColor
from ..depth120 import VarColor


class XlFont():
    """Excel 用フォント
    """


    @staticmethod
    def from_dict(contents_doc, xl_font_dict):
        """辞書を元に生成
        """
        web_safe_color_code = None
        if 'foreground' in xl_font_dict and (foreground_dict := xl_font_dict['foreground']):
            if 'varColor' in foreground_dict and (fg_color := foreground_dict['varColor']):
                var_color_obj = VarColor(fg_color)
                web_safe_color_obj = var_color_obj.to_web_safe_color_obj(
                        contents_doc=contents_doc)
                web_safe_color_code = web_safe_color_obj.code

        return XlFont(
                web_safe_color_code=web_safe_color_code)


    def __init__(self, web_safe_color_code):
        self._web_safe_color_code = web_safe_color_code


    @property
    def web_safe_color_code(self):
        return self._web_safe_color_code


    @property
    def color_code_for_xl(self):
        web_safe_color_obj = WebSafeColor(self._web_safe_color_code)
        return web_safe_color_obj.to_xl()
