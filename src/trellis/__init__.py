import os
import openpyxl as xl
from openpyxl.styles import PatternFill, Font
from openpyxl.styles.borders import Border, Side
from openpyxl.drawing.image import Image as XlImage
import json

from .compiler import AutoShadowSolver, AutoSplitPillarSolver
from .renderer import render_canvas, render_all_xl_texts, render_all_rectangles, render_all_pillar_rugs, render_all_card_shadows, render_all_terminal_shadows, render_all_line_tape_shadows, render_all_cards, render_all_terminals, render_all_line_tapes
from .renderer.ruler import render_ruler
from .share import ColorSystem


class TrellisInSrc():
    """例えば
    
    import trellis as tr

    とインポートしたとき、

    tr.render_ruler(ws, contents_doc)

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
    def compile(contents_doc, config_doc):
        """コンパイル
        """
        if 'compiler' in config_doc and (compiler_dict := config_doc['compiler']):

            # auto-split-pillar
            # -----------------
            if 'auto-split-pillar' in compiler_dict and (auto_split_pillar_dict := compiler_dict['auto-split-pillar']):
                if 'enabled' in auto_split_pillar_dict and (enabled := auto_split_pillar_dict['enabled']) and enabled:
                    # 中間ファイル（JSON形式）
                    file_path_of_contents_doc_object = auto_split_pillar_dict['objectFile']

                    print(f"""\
        auto-split-pillar
            {file_path_of_contents_doc_object=}""")


                    # ドキュメントに対して、自動ピラー分割の編集を行います
                    AutoSplitPillarSolver.edit_document(contents_doc)
                    with open(file_path_of_contents_doc_object, mode='w', encoding='utf-8') as f:
                        f.write(json.dumps(contents_doc, indent=4, ensure_ascii=False))

                    # with open(file_path_of_contents_doc_object, mode='r', encoding='utf-8') as f:
                    #     contents_doc = json.load(f)


            # auto_shadow
            # -----------
            if 'auto-shadow' in compiler_dict and (auto_shadow_dict := compiler_dict['auto-shadow']):
                if 'enabled' in auto_shadow_dict and (enabled := auto_shadow_dict['enabled']) and enabled:
                    # 中間ファイル（JSON形式）
                    file_path_of_contents_doc_object = auto_shadow_dict['objectFile']

                    print(f"""\
        auto_shadow
            {file_path_of_contents_doc_object=}""")

                    # ドキュメントに対して、影の自動設定の編集を行います
                    AutoShadowSolver.edit_document(contents_doc)

                    with open(file_path_of_contents_doc_object, mode='w', encoding='utf-8') as f:
                        f.write(json.dumps(contents_doc, indent=4, ensure_ascii=False))

                    # with open(file_path_of_contents_doc_object, mode='r', encoding='utf-8') as f:
                    #     contents_doc = json.load(f)


    @staticmethod
    def render_to_worksheet(ws, contents_doc):
        """ワークシートへの描画
        """
        # 色システムの設定
        global ColorSystem
        ColorSystem.set_color_system(ws, contents_doc)

        # キャンバスの編集
        render_canvas(ws, contents_doc)

        # 全てのテキストの描画（定規の番号除く）
        render_all_xl_texts(ws, contents_doc)

        # 全ての矩形の描画
        render_all_rectangles(ws, contents_doc)

        # 全ての柱の敷物の描画
        render_all_pillar_rugs(ws, contents_doc)

        # 全てのカードの影の描画
        render_all_card_shadows(ws, contents_doc)

        # 全ての端子の影の描画
        render_all_terminal_shadows(ws, contents_doc)

        # 全てのラインテープの影の描画
        render_all_line_tape_shadows(ws, contents_doc)

        # 全てのカードの描画
        render_all_cards(ws, contents_doc)

        # 全ての端子の描画
        render_all_terminals(ws, contents_doc)

        # 全てのラインテープの描画
        render_all_line_tapes(ws, contents_doc)

        # 定規の描画
        #       柱を上から塗りつぶすように描きます
        render_ruler(ws, contents_doc)


######################
# MARK: trellis_in_src
######################
trellis_in_src = TrellisInSrc()
