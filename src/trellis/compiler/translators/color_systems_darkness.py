from ..translator import Translator


class ColorSystemsDarkness(Translator):
    # TODO ColorSystemsDarkness


    def translate_document(self, contents_doc_rw):

        if 'colorSystem' in contents_doc_rw and (color_system_dict_rw := contents_doc_rw['colorSystem']):

            if 'darkness' in color_system_dict_rw and (darkness_dict_rw := color_system_dict_rw['darkness']):

                for key, var_color in darkness_dict_rw.items():
                    pass
