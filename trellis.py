import argparse
import datetime
import json
import os
import openpyxl as xl
import traceback

from src.trellis import trellis_in_src as tr
from src.trellis.compiler import AutoShadowSolver, AutoSplitPillarSolver


def main():
    try:
        parser = argparse.ArgumentParser()
        parser.add_argument("command", help="コマンド名")
        parser.add_argument("-c", "--config", help="設定であるJSON形式ファイルへのパス")
        parser.add_argument("-s", "--source", help="描画の指示であるJSON形式ファイルへのパス")
        parser.add_argument("-l", "--level", type=int, default=0, help="""自動化レベルです。既定値は 0。
0 で自動化は行いません。
1 で影の色の自動設定を行います。
2 で柱を跨るラインテープを自動的に別セグメントとして分割します。
""")
        parser.add_argument("-o", "--output", help="書出し先となるExcelワークブック・ファイルへのパス")
        parser.add_argument("-t", "--temp", help="テンポラリー・ディレクトリー。削除してもよいファイルを置けるディレクトリーへのパス")
        args = parser.parse_args()

        if args.command == 'init':
            canvas_width_var_value = input("""\
これからキャンバスの横幅を指定してもらいます。
よくわからないときは 100 を入力してください。
単位は［大グリッド１マス分］です。これはスプレッドシートのセル３つ分です。
例）　100
> """)
            canvas_width_obj = tr.InningsPitched(var_value=canvas_width_var_value)

            canvas_height_var_value = input("""\
これからキャンバスの縦幅を指定してもらいます。
よくわからないときは 100 を入力してください。
単位は［大グリッド１マス分］です。これはスプレッドシートのセル３つ分です。
例）　100
> """)
            canvas_height_obj = tr.InningsPitched(var_value=canvas_height_var_value)

            json_path_to_write = input("""\
これから、JSON形式ファイルの書出し先パスを指定してもらいます。
よくわからないときは ./temp/lesson/hello_world.json と入力してください、
例）　./temp/lesson/hello_world.json
# > """)
            print(f'{json_path_to_write=}')

            contents_doc = {
                "canvas": {
                    "bounds": {
                        "left": 0,
                        "top": 0,
                        "width": canvas_width_obj.var_value,
                        "height": canvas_height_obj.var_value
                    }
                },
                "ruler": {
                    "visible": True,
                    "fgColor": [
                        "xlPale.xlWhite",
                        "xlDeep.xlWhite"
                    ],
                    "bgColor": [
                        "xlDeep.xlWhite",
                        "xlPale.xlWhite"
                    ]
                }
            }

            with open(json_path_to_write, mode='w', encoding='utf-8') as f:
                f.write(json.dumps(contents_doc, indent=4, ensure_ascii=False))

            print(f"""\
{json_path_to_write} ファイルを書き出しました。確認してください。
""")


        elif args.command == 'build':
            config_doc_path_to_read = args.config   # json path
            contents_doc_path_to_read = args.source   # json path
            automation_level = args.level
            wb_path_to_write = args.output
            temporary_directory_path = args.temp

            if not config_doc_path_to_read:
                print(f"""ERROR: build コマンドには --config オプションを付けて、トレリスの設定が書かれた JSON ファイルへのパスを指定してください""")
                return

            if not contents_doc_path_to_read:
                print(f"""ERROR: build コマンドには --source オプションを付けて、描画の設定が書かれた JSON ファイルへのパスを指定してください""")
                return

            if not temporary_directory_path:
                print(f"""ERROR: build コマンドには --temp オプションを付けて、（消えても構わないファイルを入れておくための）テンポラリー・ディレクトリーのパスを指定してください""")
                return


            def get_paths(path_to_read):
                directory_path = os.path.split(path_to_read)[0]
                basename_without_ext = os.path.splitext(os.path.basename(path_to_read))[0]
                extension_with_dot = os.path.splitext(path_to_read)[1]
                print(f"""\
{directory_path=}
{basename_without_ext=}
{extension_with_dot=}
""")
                return directory_path, basename_without_ext, extension_with_dot


            config_doc_directory_path, config_doc_basename_without_ext, config_doc_extension_with_dot = get_paths(config_doc_path_to_read)
            contents_doc_directory_path, contents_doc_basename_without_ext, contents_doc_extension_with_dot = get_paths(contents_doc_path_to_read)


            # ソースファイル（JSON形式）を読込
            print(f"🔧　read {config_doc_path_to_read} file")
            with open(config_doc_path_to_read, encoding='utf-8') as f:
                config_doc = json.load(f)


            # ソースファイル（JSON形式）を読込
            print(f"🔧　read {contents_doc_path_to_read} file")
            with open(contents_doc_path_to_read, encoding='utf-8') as f:
                contents_doc = json.load(f)

            # 自動化レベル２
            if 1 < automation_level:
                # ドキュメントに対して、自動ピラー分割の編集を行います
                AutoSplitPillarSolver.edit_document(contents_doc)

                file_path_in_2_more_steps = os.path.join(temporary_directory_path, f"""{contents_doc_basename_without_ext}.in-auto-gen-2-more-steps{contents_doc_extension_with_dot}""")

                print(f"🔧　write {file_path_in_2_more_steps} file")
                with open(file_path_in_2_more_steps, mode='w', encoding='utf-8') as f:
                    f.write(json.dumps(contents_doc, indent=4, ensure_ascii=False))

                print(f"🔧　read {file_path_in_2_more_steps} file")
                with open(file_path_in_2_more_steps, mode='r', encoding='utf-8') as f:
                    contents_doc = json.load(f)

            # 自動化レベル１
            if 0 < automation_level:
                # ドキュメントに対して、影の自動設定の編集を行います
                AutoShadowSolver.edit_document(contents_doc)

                file_path_in_1_more_step = os.path.join(temporary_directory_path, f"""{contents_doc_basename_without_ext}.in-auto-gen-1-more-step{contents_doc_extension_with_dot}""")

                print(f"🔧　write {file_path_in_1_more_step} file")
                with open(file_path_in_1_more_step, mode='w', encoding='utf-8') as f:
                    f.write(json.dumps(contents_doc, indent=4, ensure_ascii=False))

                print(f"🔧　read {file_path_in_1_more_step} file")
                with open(file_path_in_1_more_step, mode='r', encoding='utf-8') as f:
                    contents_doc = json.load(f)

            # ワークブックを新規生成
            wb = xl.Workbook()

            # ワークシート
            ws = wb['Sheet']

            # ワークシートへの描画
            tr.render_to_worksheet(ws, contents_doc)

            # ワークブックの保存
            print(f"🔧　write {wb_path_to_write} file")
            wb.save(wb_path_to_write)

            print(f"Finished. Please look {wb_path_to_write} file.")

        else:
            raise ValueError(f'unsupported command: {args.command}')


    except Exception as err:
        print(f"""\
[{datetime.datetime.now()}] おお、残念！　例外が投げられてしまった！
{type(err)=}  {err=}

以下はスタックトレース表示じゃ。
{traceback.format_exc()}
""")


########################################
# コマンドから実行時
########################################
if __name__ == '__main__':
    """コマンドから実行時"""
    main()
