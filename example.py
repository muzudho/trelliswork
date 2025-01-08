"""
例を実行します

py example.py all
"""

import traceback
import datetime
import sys


########################################
# コマンドから実行時
########################################
if __name__ == '__main__':
    """コマンドから実行時"""

    try:
        args = sys.argv

        if 1 < len(args):
            if args[1] == 'all':
                from examples.step1_render_empty import *
                from examples.step2_render_pillar_header import *

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
