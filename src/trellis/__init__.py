import os
import openpyxl as xl
from openpyxl.styles import PatternFill, Font
from openpyxl.styles.borders import Border, Side
from openpyxl.drawing.image import Image as XlImage
import json

from .renderer import render_canvas, render_all_xl_texts, render_all_rectangles, render_all_pillar_rugs, render_all_card_shadows, render_all_terminal_shadows, render_all_line_tape_shadows, render_all_cards, render_all_terminals, render_all_line_tapes
from .renderer.ruler import render_ruler
from .share import ColorSystem


class TrellisInSrc():
    """例えば
    
    import trellis as tr

    とインポートしたとき、

    tr.render_ruler(ws, document)

    という形で関数を呼び出せるようにしたラッパー
    """


    @staticmethod
    def InningsPitched(var_value=None, integer_part=None, decimal_part=None):
        global InningsPitched
        if var_value:
            return InningsPitched.from_var_value(var_value)
        elif integer_part or decimal_part:
            return InningsPitched.from_integer_and_decimal_part(integer_part, decimal_part)
        else:
            raise ValueError(f'{var_value=} {integer_part=} {decimal_part=}')


    @staticmethod
    def render_to_worksheet(ws, document):
        """ワークシートへの描画
        """
        # 色システムの設定
        global ColorSystem
        ColorSystem.set_color_system(ws, document)

        # キャンバスの編集
        render_canvas(ws, document)

        # 全てのテキストの描画（定規の番号除く）
        render_all_xl_texts(ws, document)

        # 全ての矩形の描画
        render_all_rectangles(ws, document)

        # 全ての柱の敷物の描画
        render_all_pillar_rugs(ws, document)

        # 全てのカードの影の描画
        render_all_card_shadows(ws, document)

        # 全ての端子の影の描画
        render_all_terminal_shadows(ws, document)

        # 全てのラインテープの影の描画
        render_all_line_tape_shadows(ws, document)

        # 全てのカードの描画
        render_all_cards(ws, document)

        # 全ての端子の描画
        render_all_terminals(ws, document)

        # 全てのラインテープの描画
        render_all_line_tapes(ws, document)

        # 定規の描画
        #       柱を上から塗りつぶすように描きます
        render_ruler(ws, document)


######################
# MARK: trellis_in_src
######################
trellis_in_src = TrellisInSrc()
