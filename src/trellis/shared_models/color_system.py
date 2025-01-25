import re

from openpyxl.styles import PatternFill


class ColorSystem():
    """色システム
    """


    _none_pattern_fill = PatternFill(patternType=None)

    @classmethod
    @property
    def none_pattern_fill(clazz):
        return clazz._none_pattern_fill


    @classmethod
    def alias_to_web_safe_color_dict(clazz, contents_doc):
        # # TODO 前もって作っておきたい
        # if 'colorSystem' not in contents_doc:
        #     return {}
        
        # if 'alias' not in contents_doc['colorSystem']:
        #     return {}

        return contents_doc['colorSystem']['alias']


    @staticmethod
    def solve_tone_and_color_name(contents_doc, tone_and_color_name):
        try:
            tone, color = tone_and_color_name.split('.', 2)
        except:
            print(f'solve_tone_and_color_name: tone.color の形式でない {tone_and_color_name=}')
            raise


        tone = tone.strip()
        color = color.strip()

        if tone in ColorSystem.alias_to_web_safe_color_dict(contents_doc) and (tone_dict := ColorSystem.alias_to_web_safe_color_dict(contents_doc)[tone]):
            if color in tone_dict and (web_safe_color_code := tone_dict[color]):
                return web_safe_color_code

        print(f'solve_tone_and_color_name: 色がない {tone_and_color_name=}')
        return None
