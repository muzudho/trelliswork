"""
白紙に柱の頭を追加
"""

import json

from src.trellis import trellis_in_src as tr


# 設定ファイル（JSON形式）
file_path_of_config_doc = './examples/data/trellis-config-of-example2.json'

print(f"""\
example 2: pillars
    {file_path_of_config_doc=}""")

# 設定ファイル（JSON形式）を読込
with open(file_path_of_config_doc, encoding='utf-8') as f:
    config_doc = json.load(f)


# ビルド
tr.build(
        config_doc=config_doc)
