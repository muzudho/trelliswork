from .color_system import ColorSystem


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
                web_safe_color_code = ColorSystem.var_color_name_to_web_safe_color_code(
                        contents_doc=contents_doc,
                        var_color_name=fg_color)

        return XlFont(
                web_safe_color_code=web_safe_color_code)


    def __init__(self, web_safe_color_code):
        self._web_safe_color_code = web_safe_color_code


    @property
    def web_safe_color_code(self):
        return self._web_safe_color_code


    @property
    def color_code_for_xl(self):
        return ColorSystem.web_safe_color_code_to_xl(self._web_safe_color_code)
