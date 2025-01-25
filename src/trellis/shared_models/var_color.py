import re


class VarColor():
    """様々な色指定
    """


    @classmethod
    @property
    def AUTO(clazz):
        return 1


    @classmethod
    @property
    def DARKNESS(clazz):
        return 2


    @classmethod
    @property
    def PAPER_COLOR(clazz):
        return 3


    @classmethod
    @property
    def TONE_AND_COLOR_NAME(clazz):
        return 4


    @classmethod
    @property
    def WEB_SAFE_COLOR_CODE(clazz):
        return 5


    @classmethod
    @property
    def XL_COLOR_CODE(clazz):
        return 6


    @staticmethod
    def what_am_i(var_color_name):
        """トーン名・色名の欄に何が入っているか判定します
        """

        # 何も入っていない、または False が入っている
        if not var_color_name:
            return False

        # ナンが入っている
        if var_color_name is None:
            return None

        if isinstance(var_color_name, dict):
            var_color_dict = var_color_name
            if 'darkness' in var_color_dict:
                return VarColor.DARKNESS
            
            else:
                raise ValueError(f'未定義の色指定。 {var_color_name=}')


        # ウェブ・セーフ・カラーが入っている
        #
        #   とりあえず、 `#` で始まるなら、ウェブセーフカラーとして扱う
        #
        #if var_color_name.startswith('#'):
        if re.match(r'^#[0-9a-fA-f]{6}$', var_color_name):
            return VarColor.WEB_SAFE_COLOR_CODE

        if re.match(r'^[0-9a-fA-f]{6}$', var_color_name):
            return VarColor.XL_COLOR_CODE

        # 色相名と色名だ
        #if '.' in var_color_name:
        if re.match(r'^[0-9a-zA-Z_]+\.[0-9a-zA-Z_]+$', var_color_name):
            return VarColor.TONE_AND_COLOR_NAME

        # "auto", "paperColor" キーワードのいずれかが入っている
        if var_color_name in ["auto", "paperColor"]:
            return var_color_name
        
        raise ValueError(f"""ERROR: what_am_i: undefined {var_color_name=}""")


    def __init__(self, var_color_value):
        self._var_type = VarColor.what_am_i(var_color_value)


    @property
    def var_type(self):
        return self._var_type
