import os
import openpyxl as xl
import json

from .compiler import Compiler # Compiler クラスはパッケージ利用者のために公開します
from .compiler_parts import AutoShadow, AutoSplitSegmentByPillar, Imports, ResolveAliasOfColor, ResolveVarBounds
from .renderer.features import render_canvas, render_all_cards, render_all_line_tapes, render_all_pillar_rugs, render_all_rectangles, render_ruler, render_shadow_of_all_cards, render_shadow_of_all_line_tapes, render_shadow_of_all_terminals, render_all_terminals, render_all_xl_texts
from .shared_models import FilePath, InningsPitched


@staticmethod
def render_to_worksheet(config_dict, contents_dict, ws):
    """ワークシートへの描画
    """

    # キャンバスの編集
    render_canvas(
            config_doc=config_dict,
            contents_doc=contents_dict,
            ws=ws)

    # 全てのテキストの描画（定規の番号除く）
    render_all_xl_texts(
            config_doc=config_dict,
            contents_doc=contents_dict,
            ws=ws)

    # 全ての矩形の描画
    render_all_rectangles(
            config_doc=config_dict,
            contents_doc=contents_dict,
            ws=ws)

    # 全ての柱の敷物の描画
    render_all_pillar_rugs(
            config_doc=config_dict,
            contents_doc=contents_dict,
            ws=ws)

    # 全てのカードの影の描画
    render_shadow_of_all_cards(
            config_doc=config_dict,
            contents_doc=contents_dict,
            ws=ws)

    # 全ての端子の影の描画
    render_shadow_of_all_terminals(
            config_doc=config_dict,
            contents_doc=contents_dict,
            ws=ws)

    # 全てのラインテープの影の描画
    render_shadow_of_all_line_tapes(
            config_doc=config_dict,
            contents_doc=contents_dict,
            ws=ws)

    # 全てのカードの描画
    render_all_cards(
            config_doc=config_dict,
            contents_doc=contents_dict,
            ws=ws)

    # 全ての端子の描画
    render_all_terminals(
            config_doc=config_dict,
            contents_doc=contents_dict,
            ws=ws)

    # 全てのラインテープの描画
    render_all_line_tapes(
            config_doc=config_dict,
            contents_doc=contents_dict,
            ws=ws)

    # 定規の描画
    #       柱を上から塗りつぶすように描きます
    render_ruler(
            config_doc=config_dict,
            contents_doc=contents_dict,
            ws=ws)


class Trellis():
    """トレリス"""


    @staticmethod
    def init():
        """コンテンツ・ファイルを出力する
        """

        canvas_width_var_value = input("""\
これからキャンバスの横幅を指定してもらいます。
よくわからないときは 100 を入力してください。
単位は［大グリッド１マス分］です。これはスプレッドシートのセル３つ分です。
例）　100
> """)

        canvas_width_obj = InningsPitched.from_var_value(var_value=canvas_width_var_value)

        canvas_height_var_value = input("""\
これからキャンバスの縦幅を指定してもらいます。
よくわからないときは 100 を入力してください。
単位は［大グリッド１マス分］です。これはスプレッドシートのセル３つ分です。
例）　100
> """)
        canvas_height_obj = InningsPitched.from_var_value(var_value=canvas_height_var_value)

        json_path_to_write = input("""\
これから、JSON形式ファイルの書出し先パスを指定してもらいます。
よくわからないときは ./temp/lesson/hello_world.json と入力してください、
例）　./temp/lesson/hello_world.json
# > """)
        print(f'{json_path_to_write=}')

        contents_doc = {
            "imports": [
                "./examples/data_of_contents/alias_for_color.json"
            ],
            "canvas": {
                "varBounds": {
                    "left": 0,
                    "top": 0,
                    "width": canvas_width_obj.var_value,
                    "height": canvas_height_obj.var_value
                }
            },
            "ruler": {
                "visible": True,
                "foreground": {
                    "varColors": [
                        "xlPale.xlWhite",
                        "xlDeep.xlWhite"
                    ]
                },
                "background": {
                    "varColors": [
                        "xlDeep.xlWhite",
                        "xlPale.xlWhite"
                    ]
                }
            }
        }

        print(f"🔧　write {json_path_to_write} file.")
        with open(json_path_to_write, mode='w', encoding='utf-8') as f:
            f.write(json.dumps(contents_doc, indent=4, ensure_ascii=False))

        print(f"""\
{json_path_to_write} ファイルを書き出しました。確認してください。
""")


    @staticmethod
    def build(
            config,
            content,
            temp_dir,
            workbook):
        """ビルド

        Parameters
        ----------
        config : str
            コンフィグ・ファイル（読取用）へのパス。
        content : str
            コンテント・ファイル（読取用）へのパス。
        temp_dir : str
            消してもいいファイルだけが入っているディレクトリー
        workbook : str
            ワークブック（書込用）へのパス。拡張子が `.xlsx` のファイルを想定しています。
        """

        if not config:
            print(f"""ERROR: build() の config 引数には、トレリスワークの設定が書かれた JSON ファイルへのパスを指定してください""")
            return

        if not content:
            print(f"""ERROR: build() の content 引数には、描画の設定が書かれた JSON ファイルへのパスを指定してください""")
            return

        if not workbook:
            print(f"""ERROR: build() の workbook 引数には、保存先のワークブック・ファイル（.xslx）へのパスを指定してください""")
            return

        if not temp_dir:
            print(f"""ERROR: build() の temp_dir 引数には、（消えても構わないファイルを入れておくための）テンポラリー・ディレクトリーのパスを指定してください""")
            return


        # ソースファイル（JSON形式）を読込
        print(f"🔧　read {config} file.")
        with open(config, encoding='utf-8') as f:
            config_dict = json.load(f)


        # コマンドライン引数で設定を上書き
        if 'builder' not in config_dict:
            config_dict['builder'] = {}
        
        config_dict['builder']['--temp'] = temp_dir

        if 'compiler' not in config_dict:
            config_dict['compiler'] = {}
            config_dict['compiler']['source'] = content

        if 'renderer' not in config_dict:
            config_dict['renderer'] = {}

        config_dict['renderer']['--output'] = workbook


        # ビルド
        Trellis.build_by_config_doc(
                config_dict=config_dict)


    @staticmethod
    def build_by_config_doc(config_dict):
        """ビルド

        Compiler()._compile_by_dict() と render_to_worksheet() を呼び出します。
        """

        # ソースファイル（JSON形式）読込
        file_path_of_contents_doc = config_dict['compiler']['source']
        print(f"🔧　read {file_path_of_contents_doc} file.")
        with open(file_path_of_contents_doc, encoding='utf-8') as f:
            contents_dict = json.load(f)

        # 出力ファイル（JSON形式）
        wb_path_to_write = config_dict['renderer']['--output']

        # コンパイル
        tc = Compiler()
        tc._compile_by_dict(
                config=config_dict,
                content=contents_dict)

        # ワークブックを新規生成
        wb = xl.Workbook()

        # ワークシート
        ws = wb['Sheet']

        # ワークシートへの描画
        render_to_worksheet(
                config_dict=config_dict,
                contents_dict=contents_dict,
                ws=ws)

        # ワークブックの保存
        print(f"🔧　write {wb_path_to_write} file.")
        wb.save(wb_path_to_write)

        print(f"Finished. Please look {wb_path_to_write} file.")
