import argparse
import datetime
import json
import os
import openpyxl as xl
import traceback
from src.trellis import trellis_in_src as tr


def main():
    try:
        parser = argparse.ArgumentParser()
        parser.add_argument("command", help="コマンド名")
        parser.add_argument("-f", "--file", help="元となるJSON形式ファイルへのパス")
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

            document = {
                "canvas": {
                    "left": 0,
                    "top": 0,
                    "width": canvas_width_obj.var_value,
                    "height": canvas_height_obj.var_value
                },
                "ruler": {
                    "visible": True,
                    "fgColor": [
                        "xl_pale.xl_white",
                        "xl_deep.xl_white"
                    ],
                    "bgColor": [
                        "xl_deep.xl_white",
                        "xl_pale.xl_white"
                    ]
                }
            }

            with open(json_path_to_write, mode='w', encoding='utf-8') as f:
                f.write(json.dumps(document, indent=4, ensure_ascii=False))

            print(f"""\
{json_path_to_write} ファイルを書き出しました。確認してください。
""")


        elif args.command == 'compile':
            json_path_to_read = args.file
            automation_level = args.level
            wb_path_to_write = args.output
            temporary_directory_path = args.temp

            if not temporary_directory_path:
                print(f"""ERROR: compile コマンドには --temp オプションを付けて、（消えても構わないファイルを入れておくための）テンポラリー・ディレクトリーのパスを指定してください""")
                return

            source_file_directory_path = os.path.split(json_path_to_read)[0]
            source_file_basename_without_ext = os.path.splitext(os.path.basename(json_path_to_read))[0]
            source_file_extension_with_dot = os.path.splitext(json_path_to_read)[1]
            print(f"""\
{source_file_directory_path=}
{source_file_basename_without_ext=}
{source_file_extension_with_dot=}
""")

            # ソースファイル（JSON形式）を読込
            print(f"🔧　read {json_path_to_read} file")
            with open(json_path_to_read, encoding='utf-8') as f:
                document = json.load(f)

            # 自動化レベル２
            if 1 < automation_level:
                # ドキュメントに対して、自動ピラー分割の編集を行います
                tr.edit_document_and_solve_auto_split_pillar(document)

                file_path_in_2_more_steps = os.path.join(temporary_directory_path, f"""{source_file_basename_without_ext}.in-auto-gen-2-more-steps{source_file_extension_with_dot}""")

                print(f"🔧　write {file_path_in_2_more_steps} file")
                with open(file_path_in_2_more_steps, mode='w', encoding='utf-8') as f:
                    f.write(json.dumps(document, indent=4, ensure_ascii=False))

                print(f"🔧　read {file_path_in_2_more_steps} file")
                with open(file_path_in_2_more_steps, mode='r', encoding='utf-8') as f:
                    document = json.load(f)

            # 自動化レベル１
            if 0 < automation_level:
                # ドキュメントに対して、影の自動設定の編集を行います
                tr.edit_document_and_solve_auto_shadow(document)

                file_path_in_1_more_step = os.path.join(temporary_directory_path, f"""{source_file_basename_without_ext}.in-auto-gen-1-more-step{source_file_extension_with_dot}""")

                print(f"🔧　write {file_path_in_1_more_step} file")
                with open(file_path_in_1_more_step, mode='w', encoding='utf-8') as f:
                    f.write(json.dumps(document, indent=4, ensure_ascii=False))

                print(f"🔧　read {file_path_in_1_more_step} file")
                with open(file_path_in_1_more_step, mode='r', encoding='utf-8') as f:
                    document = json.load(f)

            # ワークブックを新規生成
            wb = xl.Workbook()

            # ワークシート
            ws = wb['Sheet']

            # キャンバスの編集
            tr.edit_canvas(ws, document)

            # 全てのテキストの描画（定規の番号除く）
            tr.render_all_xl_texts(ws, document)

            # 全ての矩形の描画
            tr.render_all_rectangles(ws, document)

            # 全ての矩形の描画
            tr.render_all_rectangles(ws, document)

            # 全ての柱の敷物の描画
            tr.render_all_pillar_rugs(ws, document)

            # 全てのカードの影の描画
            tr.render_all_card_shadows(ws, document)

            # 全ての端子の影の描画
            tr.render_all_terminal_shadows(ws, document)

            # 全てのラインテープの影の描画
            tr.render_all_line_tape_shadows(ws, document)

            # 全てのカードの描画
            tr.render_all_cards(ws, document)

            # 全ての端子の描画
            tr.render_all_terminals(ws, document)

            # 全てのラインテープの描画
            tr.render_all_line_tapes(ws, document)

            # 定規の描画
            #       柱を上から塗りつぶすように描きます
            tr.render_ruler(ws, document)

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
