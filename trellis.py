import argparse
import datetime
import json
import os
import openpyxl as xl
import traceback
from src.trellis import trellis_in_src as tr


########################################
# コマンドから実行時
########################################
if __name__ == '__main__':
    """コマンドから実行時"""

    try:
        parser = argparse.ArgumentParser()
        parser.add_argument("command", help="コマンド名")
        parser.add_argument("-f", "--file", help="元となるJSON形式ファイルへのパス")
        parser.add_argument("-o", "--output", help="書出し先となるExcelワークブック・ファイルへのパス")
        args = parser.parse_args()

        if args.command == 'init':
            canvas_width = input("""\
これからキャンバスの横幅を指定してもらいます。
よくわからないときは 100 を入力してください。
単位は［グリッド大１マス］です。
例）　100
> """)
            canvas_width = int(canvas_width)

            canvas_height = input("""\
これからキャンバスの縦幅を指定してもらいます。
よくわからないときは 100 を入力してください。
単位は［グリッド大１マス］です。
例）　100
> """)
            canvas_height = int(canvas_height)

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
                    "width": canvas_width,
                    "height": canvas_height
                }
            }

            with open(json_path_to_write, mode='w', encoding='utf-8') as f:
                f.write(json.dumps(document, indent=4, ensure_ascii=False))

            print(f"""\
{json_path_to_write} ファイルを書き出しました。確認してください。
""")

        elif args.command == 'ruler':
            json_path_to_read = args.file
            wb_path_to_write = args.output

            print(f"🔧　read {json_path_to_read} file")
            with open(json_path_to_read, encoding='utf-8') as f:
                document = json.load(f)

            canvas_width = document['canvas']['width']
            canvas_height = document['canvas']['height']

            print(f"""{json_path_to_read} ファイルには、キャンバスの横幅 {canvas_width}、縦幅 {canvas_height} と書いてあったので、それに従って定規を描きます""")

            # ワークブックを新規生成
            wb = xl.Workbook()

            # ワークシート
            ws = wb['Sheet']

            # 定規の描画
            tr.render_ruler(document, ws)

            # ワークブックの保存
            print(f"🔧　write {wb_path_to_write} file")
            wb.save(wb_path_to_write)

            print(f"""\
{wb_path_to_write} ファイルを書き出しました。確認してください。
""")

        elif args.command == 'compile':
            json_path_to_read = args.file
            wb_path_to_write = args.output

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

            # ドキュメントに対して、自動ピラー分割の編集を行います
            tr.edit_document_and_solve_auto_split_pillar(document)

            file_path_in_2_more_steps = os.path.join(source_file_directory_path, f"""{source_file_basename_without_ext}.in-auto-gen-2-more-steps{source_file_extension_with_dot}""")

            print(f"🔧　write {file_path_in_2_more_steps} file")
            with open(file_path_in_2_more_steps, mode='w', encoding='utf-8') as f:
                f.write(json.dumps(document, indent=4, ensure_ascii=False))

            print(f"🔧　read {file_path_in_2_more_steps} file")
            with open(file_path_in_2_more_steps, mode='r', encoding='utf-8') as f:
                document = json.load(f)

            # ドキュメントに対して、影の自動設定の編集を行います
            tr.edit_document_and_solve_auto_shadow(document)

            file_path_in_1_more_step = os.path.join(source_file_directory_path, f"""{source_file_basename_without_ext}.in-auto-gen-1-more-step{source_file_extension_with_dot}""")

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

            # 全ての柱の敷物の描画
            tr.render_all_pillar_rugs(document, ws)

            # 全てのカードの影の描画
            tr.render_all_card_shadows(document, ws)

            # 全ての端子の影の描画
            tr.render_all_terminal_shadows(document, ws)

            # 全てのラインテープの影の描画
            tr.render_all_line_tape_shadows(document, ws)

            # 全てのカードの描画
            tr.render_all_cards(document, ws)

            # 全ての端子の描画
            tr.render_all_terminals(document, ws)

            # 全てのラインテープの描画
            tr.render_all_line_tapes(document, ws)

            # 定規の描画
            #       柱を上から塗りつぶすように描きます
            tr.render_ruler(document, ws)

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
