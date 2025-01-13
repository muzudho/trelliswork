import traceback
import datetime
import sys
import json


########################################
# コマンドから実行時
########################################
if __name__ == '__main__':
    """コマンドから実行時"""

    try:
        args = sys.argv

        if 1 < len(args):

            if args[1] == 'init':
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

#                 wb_path_to_read = input("""\
# 読込むワークブックへのファイルパスを指定してください。
# 例）　./temp/lesson/no1.xlsx
# > """)
#                 print(f'{wb_path_to_read=}')

#                 wb_path_to_write = input("""\
# 書出し先のワークブックへのファイルパスを指定してください。
# （読込むワークブックとは別のファイルにしてください）
# 例）　./temp/lesson/no1_2.xlsx
# > """)
#                 print(f'{wb_path_to_write=}')

            else:
                raise ValueError(f'unsupported {args[1]=}')
        
        else:
            raise ValueError(f'unsupported {len(args)=}')

    except Exception as err:
        print(f"""\
[{datetime.datetime.now()}] おお、残念！　例外が投げられてしまった！
{type(err)=}  {err=}

以下はスタックトレース表示じゃ。
{traceback.format_exc()}
""")
