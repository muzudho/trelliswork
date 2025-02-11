import os
import openpyxl as xl
import json

from .compiler.translators import AutoShadow, AutoSplitSegmentByPillar, Imports, ResolveAliasOfColor, ResolveVarBounds
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

        print(f"🔧　write {json_path_to_write} file")
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
        print(f"🔧　read {config} file")
        with open(config, encoding='utf-8') as f:
            config_dict = json.load(f)


        # コマンドライン引数で設定を上書き
        if 'builder' not in config_dict:
            config_dict['builder'] = {}
        
        config_dict['builder']['--source'] = content
        config_dict['builder']['--temp'] = temp_dir

        if 'compiler' not in config_dict:
            config_dict['compiler'] = {}

        if 'renderer' not in config_dict:
            config_dict['renderer'] = {}

        config_dict['renderer']['--output'] = workbook


        # ビルド
        Trellis.build_by_config_doc(
                config_dict=config_dict)


    @staticmethod
    def build_by_config_doc(config_dict):
        """ビルド

        Trellis.compile と render_to_worksheet を呼び出します。
        """

        # ソースファイル（JSON形式）読込
        file_path_of_contents_doc = config_dict['builder']['--source']
        print(f"🔧　read {file_path_of_contents_doc} file")
        with open(file_path_of_contents_doc, encoding='utf-8') as f:
            contents_dict = json.load(f)

        # 出力ファイル（JSON形式）
        wb_path_to_write = config_dict['renderer']['--output']

        # コンパイル
        Trellis.compile_by_dict(
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
        print(f"🔧　write {wb_path_to_write} file")
        wb.save(wb_path_to_write)

        print(f"Finished. Please look {wb_path_to_write} file.")


    @staticmethod
    def compile(config, source):
        """コンパイル

        出力ファイルは、一時ファイルという形で出力される。  

        Parameters
        ----------
        config : str
            設定ファイル（読取専用）へのパス
        source : str
            内容ファイル（読取専用）へのパス
        """

        print(f"🔧　read {config} config file")
        with open(config, encoding='utf-8') as f:
            config_dict = json.load(f)

        if 'builder' not in config_dict:
            config_dict['builder'] = {}

        if '--source' not in config_dict['builder']:
            config_dict['builder']['--source'] = source

        print(f"🔧　read {source} source file")
        with open(source, encoding='utf-8') as f:
            source_dict_rw = json.load(f)

        Trellis.compile_by_dict(
                config_dict=config_dict,
                source_dict_rw=source_dict_rw)


    @staticmethod
    def compile_by_dict(config_dict, source_dict_rw):
        """コンパイル
        TODO 出力ファイルも指定したい

        Parameters
        ----------
        config_dict : dict
            設定
        source_dict_rw : dict
            読み書き両用
        """

        source_fp = FilePath(config_dict['builder']['--source'])

        if 'compiler' in config_dict and (compiler_dict := config_dict['compiler']):

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
                            f.write(json.dumps(source_dict_rw, indent=4, ensure_ascii=False))


                # ［翻訳者一覧］
                translator_object_dict = {
                    'autoSplitSegmentByPillar': AutoSplitSegmentByPillar(),
                    'autoShadow': AutoShadow(),
                    'imports': Imports(),
                    'resolveAliasOfColor': ResolveAliasOfColor(),
                    'resolveVarBounds': ResolveVarBounds(),
                }

                # 翻訳の実行順序
                if 'translationOrder' in compiler_dict and (translation_order_list := compiler_dict['translationOrder']):

                    for translation_key in translation_order_list:

                        # 各［翻訳者］
                        #
                        #   翻訳者は translate_document(source_dict_rw) というインスタンス・メソッドを持つ
                        #
                        translator_dict = translators_dict[translation_key]

                        if translation_key in translator_object_dict:
                            translator_obj = translator_object_dict[translation_key]

                            if 'enabled' in translator_dict and (enabled := translator_dict['enabled']) and enabled:
                                # ドキュメントに対して、自動ピラー分割の編集を行います
                                translator_obj.translate_document(
                                        contents_dict_rw=source_dict_rw)

                            # （場合により）中間ファイルの書出し
                            write_object_file(comment=translation_key)
