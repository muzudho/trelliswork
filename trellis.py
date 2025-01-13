import traceback
import datetime
import json
import argparse
from src.trellis import trellis_in_src as tr
import openpyxl as xl


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
            wb.save(wb_path_to_write)

            print(f"""\
{wb_path_to_write} ファイルを書き出しました。確認してください。
""")

        else:
            raise ValueError(f'unsupported command: {args.command}')


    except Exception as err:
        print(f"""\
[{datetime.datetime.now()}] おお、残念！　例外が投げられてしまった！
{type(err)=}  {err=}

以下はスタックトレース表示じゃ。
{traceback.format_exc()}
""")
