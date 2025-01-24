import os
import openpyxl as xl
from openpyxl.styles import PatternFill, Font
from openpyxl.styles.borders import Border, Side
from openpyxl.drawing.image import Image as XlImage
import json

from .compiler.translators.auto_shadow import AutoShadow
from .compiler.translators.auto_split_pillar import AutoSplitSegmentByPillar
from .compiler.translators.color_systems_darkness import ColorSystemsDarkness

from .renderer.features.canvas import render_canvas
from .renderer.features.cards import render_all_cards
from .renderer.features.line_tapes import render_all_line_tapes
from .renderer.features.pillars import render_all_pillar_rugs
from .renderer.features.rectangles import render_all_rectangles
from .renderer.features.ruler import render_ruler
from .renderer.features.shadow_of_cards import render_shadow_of_all_cards
from .renderer.features.shadow_of_line_tapes import render_shadow_of_all_line_tapes
from .renderer.features.shadow_of_terminals import render_shadow_of_all_terminals
from .renderer.features.terminals import render_all_terminals
from .renderer.features.xl_text import render_all_xl_texts
from .share import ColorSystem, FilePath


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

        source_fp = FilePath(config_doc['builder']['--source'])

        if 'compiler' in config_doc and (compiler_dict := config_doc['compiler']):

            def get_object_folder():
                if 'objectFolder' not in compiler_dict:
                    raise ValueError("""設定ファイルでコンパイラーの処理結果を中間ファイルとして出力する設定にした場合は、['compiler']['objectFolder']が必要です。""")

                return compiler_dict['objectFolder']


            if 'objectFilePrefix' in compiler_dict and (object_file_prefix := compiler_dict['objectFilePrefix']) and object_file_prefix is not None:
                pass
            else:
                object_file_prefix = ''


            if 'tlanslators' in compiler_dict and (translators_dict := compiler_dict['tlanslators']):


                def create_file_path_of_contents_doc_object(source_fp, object_file_dict):
                    """中間ファイルのパス作成"""
                    object_suffix = object_file_dict['suffix']
                    basename = f'{object_file_prefix}__{source_fp.basename_without_ext}__{object_suffix}.json'
                    return os.path.join(get_object_folder(), basename)


                def write_object_file(comment):
                    """中間ファイルの書出し
                    """
                    if 'objectFile' in translator_dict and (object_file_dict := translator_dict['objectFile']):
                        file_path_of_contents_doc_object = create_file_path_of_contents_doc_object(
                                source_fp=source_fp,
                                object_file_dict=object_file_dict)

                        print(f"""\
🔧　write {file_path_of_contents_doc_object} file
    {comment}""")

                        # ディレクトリーが存在しなければ作成する
                        directory_path = os.path.split(file_path_of_contents_doc_object)[0]
                        os.makedirs(directory_path, exist_ok=True)

                        print(f"🔧　write {file_path_of_contents_doc_object} file")
                        with open(file_path_of_contents_doc_object, mode='w', encoding='utf-8') as f:
                            f.write(json.dumps(contents_doc_rw, indent=4, ensure_ascii=False))


                # ［翻訳者一覧］
                translator_object_dict = {
                    'autoSplitSegmentByPillar': AutoSplitSegmentByPillar(),
                    'autoShadow': AutoShadow(),
                }

                # 各［翻訳者］
                #
                #   翻訳者は translate_document(contents_doc_rw) というインスタンス・メソッドを持つ
                #
                for key, translator_dict in translators_dict.items():
                    if key in translator_object_dict:
                        translator_obj = translator_object_dict[key]

                        if 'enabled' in translator_dict and (enabled := translator_dict['enabled']) and enabled:
                            # ドキュメントに対して、自動ピラー分割の編集を行います
                            translator_obj.translate_document(
                                    contents_doc_rw=contents_doc_rw)

                        # （場合により）中間ファイルの書出し
                        write_object_file(comment=key)


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
        render_shadow_of_all_cards(
                config_doc=config_doc,
                contents_doc=contents_doc,
                ws=ws)

        # 全ての端子の影の描画
        render_shadow_of_all_terminals(
                config_doc=config_doc,
                contents_doc=contents_doc,
                ws=ws)

        # 全てのラインテープの影の描画
        render_shadow_of_all_line_tapes(
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
