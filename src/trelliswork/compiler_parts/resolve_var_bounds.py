import copy

from ..shared_models import InningsPitched
from . import Part


class ResolveVarBounds(Part):
    """varBounds を bounds へ翻訳します
    """


    def compile_document(self, contents_dict_rw):

        # 再帰的に更新
        ResolveVarBounds.search_dict(
                contents_dict_rw=contents_dict_rw,
                current_dict_rw=contents_dict_rw)


    @staticmethod
    def search_dict(contents_dict_rw, current_dict_rw):

        new_bounds_dict = None

        for key, value in current_dict_rw.items():
            if key == 'varBounds':

                # 辞書 varColorDict の辞書要素
                if isinstance(value, dict):
                    new_bounds_dict = ResolveVarBounds.search_var_bounds_dict(
                            contents_dict_rw=contents_dict_rw,
                            current_var_bounds_dict_rw=value)
                    continue


            # 辞書の任意のキーの辞書要素
            if isinstance(value, dict):
                ResolveVarBounds.search_dict(
                        contents_dict_rw=contents_dict_rw,
                        current_dict_rw=value)

            # 辞書の任意のキーのリスト要素
            elif isinstance(value, list):
                ResolveVarBounds.search_list(
                        contents_dict_rw=contents_dict_rw,
                        current_list_rw=value)

        # 更新した項目を追加
        if new_bounds_dict:
            del current_dict_rw['varBounds']
            current_dict_rw['bounds'] = new_bounds_dict


    @staticmethod
    def search_list(contents_dict_rw, current_list_rw):
        for index, value in enumerate(current_list_rw):

            # リストの辞書要素
            if isinstance(value, dict):
                ResolveVarBounds.search_dict(
                        contents_dict_rw=contents_dict_rw,
                        current_dict_rw=value)

            # リストのリスト要素
            elif isinstance(value, list):
                ResolveVarBounds.search_list(
                        contents_dict_rw=contents_dict_rw,
                        current_list_rw=value)

    @staticmethod
    def search_var_bounds_dict(contents_dict_rw, current_var_bounds_dict_rw):
        """left, top, right, bottom, width, height が［投球回］形式の辞書
        """
        bounds_dict = copy.deepcopy(current_var_bounds_dict_rw)

        if 'left' in bounds_dict:
            bounds_obj = InningsPitched.from_var_value(bounds_dict['left'])
            bounds_dict['left'] = bounds_obj.total_of_out_counts_qty

        if 'top' in bounds_dict:
            bounds_obj = InningsPitched.from_var_value(bounds_dict['top'])
            bounds_dict['top'] = bounds_obj.total_of_out_counts_qty

        if 'width' in bounds_dict:
            bounds_obj = InningsPitched.from_var_value(bounds_dict['width'])
            bounds_dict['width'] = bounds_obj.total_of_out_counts_qty

        if 'height' in bounds_dict:
            bounds_obj = InningsPitched.from_var_value(bounds_dict['height'])
            bounds_dict['height'] = bounds_obj.total_of_out_counts_qty

        if 'right' in bounds_dict:
            bounds_obj = InningsPitched.from_var_value(bounds_dict['right'])
            bounds_dict['right'] = bounds_obj.total_of_out_counts_qty

        if 'bottom' in bounds_dict:
            bounds_obj = InningsPitched.from_var_value(bounds_dict['bottom'])
            bounds_dict['bottom'] = bounds_obj.total_of_out_counts_qty

        return bounds_dict