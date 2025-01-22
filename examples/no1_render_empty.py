"""
白紙の作成
"""

import json
import openpyxl as xl
from openpyxl.styles import PatternFill, Font
from src.trellis import trellis_in_src as tr


# 設定ファイル（JSON形式）
file_path_of_config_doc = './examples/data/trellis-config-of-example1.json'

print(f"""\
example 1: render empty
    {file_path_of_config_doc=}""")

# 設定ファイル（JSON形式）を読込
with open(file_path_of_config_doc, encoding='utf-8') as f:
    config_doc = json.load(f)


# ソースファイル（JSON形式）
file_path_of_contents_doc = './examples/data/battle_sequence_of_unfair_cointoss.step1_full_manual.json'
# 出力ファイル（JSON形式）
file_path_of_output = config_doc['compiler']['--output']

print(f"""\
    {file_path_of_contents_doc=}
    {file_path_of_output=}""")

# ソースファイル（JSON形式）を読込
with open(file_path_of_contents_doc, encoding='utf-8') as f:
    contents_doc = json.load(f)


# ワークブックを新規生成
wb = xl.Workbook()

# ワークシート
ws = wb['Sheet']

# ワークシートへの描画
tr.render_to_worksheet(ws, contents_doc)

# ワークブックの保存            
wb.save(file_path_of_output)
