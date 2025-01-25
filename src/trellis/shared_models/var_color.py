class VarColor():
    """様々な色指定
    """


    @classmethod
    @property
    def AUTO(clazz):
        return 'auto'


    @classmethod
    @property
    def DARKNESS(clazz):
        return 'darkness'


    @classmethod
    @property
    def PAPER_COLOR(clazz):
        return 'paperColor'


    @classmethod
    @property
    def TONE_AND_COLOR_NAME(clazz):
        return 'toneAndColorName'


    @classmethod
    @property
    def WEB_SAFE_COLOR_CODE(clazz):
        return 'webSafeColorCode'


    @classmethod
    @property
    def XL_COLOR_CODE(clazz):
        return 'xlColorCode'


    @staticmethod
    def from_var_color_value(var_color_value):
        pass
