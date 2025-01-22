import os
import openpyxl as xl
from openpyxl.styles import PatternFill, Font
from openpyxl.styles.borders import Border, Side
from openpyxl.drawing.image import Image as XlImage
import json

from .compiler.auto_shadow import AutoShadowSolver
from .compiler.auto_split_pillar import AutoSplitSegmentByPillarSolver
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
    def build(config_doc):
        """ビルド
        """

        # ソースファイル（JSON形式）読込
        file_path_of_contents_doc = config_doc['builder']['--source']
        print(f"🔧　read {file_path_of_contents_doc} file")
        with open(file_path_of_contents_doc, encoding='utf-8') as f:
            contents_doc = json.load(f)

        # 出力ファイル（JSON形式）
        wb_path_to_write = config_doc['renderer']['--output']

        # コンパイル
        TrellisInSrc.compile(
                contents_doc_rw=contents_doc,
                config_doc=config_doc)

        # ワークブックを新規生成
        wb = xl.Workbook()

        # ワークシート
        ws = wb['Sheet']

        # ワークシートへの描画
        TrellisInSrc.render_to_worksheet(
                config_doc=config_doc,
                contents_doc=contents_doc,
                ws=ws)

        # ワークブックの保存
        print(f"🔧　write {wb_path_to_write} file")
        wb.save(wb_path_to_write)

        print(f"Finished. Please look {wb_path_to_write} file.")


    @staticmethod
    def compile(contents_doc_rw, config_doc):
        """コンパイル

        Parameters
        ----------
        contents_doc_rw : dict
            読み書き両用
        """
        if 'compiler' in config_doc and (compiler_dict := config_doc['compiler']):

            # autoSplitSegmentByPillar
            # ------------------------
            if 'autoSplitSegmentByPillar' in compiler_dict and (auto_split_segment_by_pillar_dict := compiler_dict['autoSplitSegmentByPillar']):
                if 'enabled' in auto_split_segment_by_pillar_dict and (enabled := auto_split_segment_by_pillar_dict['enabled']) and enabled:
                    # 中間ファイル（JSON形式）
                    file_path_of_contents_doc_object = auto_split_segment_by_pillar_dict['objectFile']

                    print(f"""\
        🔧　write {file_path_of_contents_doc_object} file
            autoSplitSegmentByPillar""")

                    # ドキュメントに対して、自動ピラー分割の編集を行います
                    AutoSplitSegmentByPillarSolver.edit_document(
                                contents_doc_rw=contents_doc_rw)

                    # ディレクトリーが存在しなければ作成する
                    directory_path = os.path.split(file_path_of_contents_doc_object)[0]
                    os.makedirs(directory_path, exist_ok=True)

                    print(f"🔧　write {file_path_of_contents_doc_object} file")
                    with open(file_path_of_contents_doc_object, mode='w', encoding='utf-8') as f:
                        f.write(json.dumps(contents_doc_rw, indent=4, ensure_ascii=False))


            # autoShadow
            # ----------
            if 'autoShadow' in compiler_dict and (auto_shadow_dict := compiler_dict['autoShadow']):
                if 'enabled' in auto_shadow_dict and (enabled := auto_shadow_dict['enabled']) and enabled:
                    # 中間ファイル（JSON形式）
                    file_path_of_contents_doc_object = auto_shadow_dict['objectFile']

                    print(f"""\
        🔧　write {file_path_of_contents_doc_object} file
            auto_shadow""")

                    # ドキュメントに対して、影の自動設定の編集を行います
                    AutoShadowSolver.edit_document(
                                contents_doc_rw=contents_doc_rw)

                    # ディレクトリーが存在しなければ作成する
                    directory_path = os.path.split(file_path_of_contents_doc_object)[0]
                    os.makedirs(directory_path, exist_ok=True)

                    print(f"🔧　write {file_path_of_contents_doc_object} file")
                    with open(file_path_of_contents_doc_object, mode='w', encoding='utf-8') as f:
                        f.write(json.dumps(contents_doc_rw, indent=4, ensure_ascii=False))


    @staticmethod
    def render_to_worksheet(config_doc, contents_doc, ws):
        """ワークシートへの描画
        """
        # 色システムの設定
        global ColorSystem
        ColorSystem.set_color_system(
                config_doc=config_doc,
                contents_doc=contents_doc,
                ws=ws)

        # キャンバスの編集
        render_canvas(
                config_doc=config_doc,
                contents_doc=contents_doc,
                ws=ws)

        # 全てのテキストの描画（定規の番号除く）
        render_all_xl_texts(
                config_doc=config_doc,
                contents_doc=contents_doc,
                ws=ws)

        # 全ての矩形の描画
        render_all_rectangles(
                config_doc=config_doc,
                contents_doc=contents_doc,
                ws=ws)

        # 全ての柱の敷物の描画
        render_all_pillar_rugs(
                config_doc=config_doc,
                contents_doc=contents_doc,
                ws=ws)

        # 全てのカードの影の描画
        render_all_card_shadows(
                config_doc=config_doc,
                contents_doc=contents_doc,
                ws=ws)

        # 全ての端子の影の描画
        render_all_terminal_shadows(
                config_doc=config_doc,
                contents_doc=contents_doc,
                ws=ws)

        # 全てのラインテープの影の描画
        render_all_line_tape_shadows(
                config_doc=config_doc,
                contents_doc=contents_doc,
                ws=ws)

        # 全てのカードの描画
        render_all_cards(
                config_doc=config_doc,
                contents_doc=contents_doc,
                ws=ws)

        # 全ての端子の描画
        render_all_terminals(
                config_doc=config_doc,
                contents_doc=contents_doc,
                ws=ws)

        # 全てのラインテープの描画
        render_all_line_tapes(
                config_doc=config_doc,
                contents_doc=contents_doc,
                ws=ws)

        # 定規の描画
        #       柱を上から塗りつぶすように描きます
        render_ruler(
                config_doc=config_doc,
                contents_doc=contents_doc,
                ws=ws)


######################
# MARK: trellis_in_src
######################
trellis_in_src = TrellisInSrc()
