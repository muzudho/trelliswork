from ..translator import Translator


class ColorSystemsDarkness(Translator):


    def translate_document(self, contents_doc_rw):

        if 'colorSystems' in contents_doc_rw and (color_systems_dict_rw := contents_doc_rw['colorSystems']):

            if 'darkness' in color_systems_dict_rw and (darkness_dict_rw := color_systems_dict_rw['darkness']):

                for key, var_color in darkness_dict_rw.items():
                    pass
