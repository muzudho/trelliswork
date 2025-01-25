from ...shared_models import ColorSystem

from ..translator import Translator


class ResolveAliasOfColor(Translator):
    """［影色の対応表］が色の別名で指定されていれば、ウェブ・セーフ・カラー・コードに翻訳します
    """


    def translate_document(self, contents_doc_rw):

        if 'colorSystem' in contents_doc_rw and (color_system_dict_rw := contents_doc_rw['colorSystem']):

            # 別名の対応表
            # alias_dict_rw
            if 'alias' in color_system_dict_rw and (alias_dict_rw := color_system_dict_rw['alias']):
                pass

            else:
                return


            # 再帰的に更新
            ResolveAliasOfColor.search_dict(
                    contents_doc_rw=contents_doc_rw,
                    current_dict_rw=contents_doc_rw)


            new_dict = {}
            delete_keys = []


            # ［影色の対応表］
            if 'shadowColorMappings' in color_system_dict_rw and (shadow_color_mappings_dict_rw := color_system_dict_rw['shadowColorMappings']):

                # key も value も var_color_name 形式
                for key_vcn, value_vcn in shadow_color_mappings_dict_rw.items():

                    color_type = ColorSystem.what_is_var_color_name(
                            var_color_name=key_vcn)


                    if color_type == ColorSystem.TONE_AND_COLOR_NAME:
                        key_web_safe_color_code = ColorSystem.solve_tone_and_color_name(
                                contents_doc=contents_doc_rw,
                                tone_and_color_name=key_vcn)


                    # ［ウェブ・セーフ・カラー］、［紙の色］はそのまま
                    elif color_type in [ColorSystem.WEB_SAFE_COLOR_CODE, ColorSystem.PAPER_COLOR]:
                        key_web_safe_color_code = key_vcn


                    else:
                        print(f'NOT_IMPLEMENTED: ResoluveAliasOfColor: ★未実装です。 {color_type=}')
                        continue


                    color_type = ColorSystem.what_is_var_color_name(
                            var_color_name=value_vcn)

                    #print(f'★ {key_web_safe_color_code=} {value_vcn=} {color_type=}')

                    if color_type == ColorSystem.TONE_AND_COLOR_NAME:
                        value_web_safe_color_code = ColorSystem.solve_tone_and_color_name(
                                contents_doc=contents_doc_rw,
                                tone_and_color_name=value_vcn)


                    # ［ウェブ・セーフ・カラー］、［紙の色］はそのまま
                    elif color_type in [ColorSystem.WEB_SAFE_COLOR_CODE, ColorSystem.PAPER_COLOR]:
                        value_web_safe_color_code = value_vcn


                    else:
                        print(f'NOT_IMPLEMENTED: ResoluveAliasOfColor: 未実装です。 {color_type=}')
                        continue


                    # 変更される要素のキー名を記憶
                    delete_keys.append(key_vcn)

                    #print(f'★翻訳 {key_web_safe_color_code=} {value_vcn=} {value_web_safe_color_code=}')
                    new_dict[key_web_safe_color_code] = value_web_safe_color_code


            # 変更された要素を削除
            for delete_key in delete_keys:
                print(f'★キー名が変わる要素を削除 {delete_key=}')
                del shadow_color_mappings_dict_rw[delete_key]


            # 更新分を追加
            for key, value in new_dict.items():
                if key in shadow_color_mappings_dict_rw:
                    # paperColor とか
                    print(f"""ERROR: ResoluveAliasOfColor: shadowColorMappings 辞書のキーが重複しています。 {key=}""")
                    continue

                shadow_color_mappings_dict_rw[key] = value


            # # TODO 別名の対応表の削除
            # del color_system_dict_rw['alias']


    @staticmethod
    def search_dict(contents_doc_rw, current_dict_rw):
        for key, value in current_dict_rw.items():
            if key == "varColor":

                # 辞書 varColor の文字列要素
                if isinstance(value, str):
                    color_type = ColorSystem.what_is_var_color_name(
                            var_color_name=value)

                    if color_type == ColorSystem.TONE_AND_COLOR_NAME:
                        web_safe_color_code = ColorSystem.solve_tone_and_color_name(
                                contents_doc=contents_doc_rw,
                                tone_and_color_name=value)

                        current_dict_rw[key] = web_safe_color_code
                
                continue

            elif key == 'varColors':

                # 辞書 varColors の配列要素
                if isinstance(value, list):
                    ResolveAliasOfColor.search_list(
                            contents_doc_rw=contents_doc_rw,
                            current_list_rw=value)

                continue

            # 辞書の任意のキーの辞書要素
            if isinstance(value, dict):
                ResolveAliasOfColor.search_dict(
                        contents_doc_rw=contents_doc_rw,
                        current_dict_rw=value)

            # 辞書の任意のキーのリスト要素
            elif isinstance(value, list):
                ResolveAliasOfColor.search_list(
                        contents_doc_rw=contents_doc_rw,
                        current_list_rw=value)


    @staticmethod
    def search_list(contents_doc_rw, current_list_rw):
        for index, value in enumerate(current_list_rw):

            # リストの文字列要素
            if isinstance(value, str):
                color_type = ColorSystem.what_is_var_color_name(
                        var_color_name=value)

                if color_type == ColorSystem.TONE_AND_COLOR_NAME:
                    web_safe_color_code = ColorSystem.solve_tone_and_color_name(
                            contents_doc=contents_doc_rw,
                            tone_and_color_name=value)

                    current_list_rw[index] = web_safe_color_code

            # リストの辞書要素
            elif isinstance(value, dict):
                ResolveAliasOfColor.search_dict(
                        contents_doc_rw=contents_doc_rw,
                        current_dict_rw=value)

            # リストのリスト要素
            elif isinstance(value, list):
                ResolveAliasOfColor.search_list(
                        contents_doc_rw=contents_doc_rw,
                        current_list_rw=value)
