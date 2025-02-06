class WebSafeColor():
    """ウェブ・セーフ・カラー"""


    def __init__(self, code):

        # if self.var_type != VarColor.WEB_SAFE_COLOR:
        #     raise ValueError(f'web_safe_color_code_to_xl: ウェブ・セーフ・カラーじゃない。 {self.var_color_value=}')

        self._code = code


    @property
    def code(self):
        return self._code


    def to_xl(self):
        """頭の `#` を外します
        """
        return self._code[1:]
