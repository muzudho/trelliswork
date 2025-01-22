"""
白紙に柱の頭を追加
"""

import json
import openpyxl as xl
from openpyxl.styles import PatternFill, Font

from src.trellis import trellis_in_src as tr
from src.trellis.compiler import AutoShadowSolver


# 設定ファイル（JSON形式）
file_path_of_config_doc = './examples/data/trellis-config-of-example4.json'

print(f"""\
example 4: auto shadow
    {file_path_of_config_doc=}""")

# 設定ファイル（JSON形式）を読込
with open(file_path_of_config_doc, encoding='utf-8') as f:
    config_doc = json.load(f)


# ソースファイル（JSON形式）
file_path_of_contents_doc = './examples/data/battle_sequence_of_unfair_cointoss.step4_auto_shadow.json'
# オブジェクトファイル（JSON形式）
file_path_of_contents_doc_2 = './temp/examples/data_step4_battle_sequence_of_unfair_cointoss.step4_auto_shadow.compiled.json'
# 出力ファイル（JSON形式）
file_path_of_output = config_doc['compiler']['--output']

print(f"""\
    {file_path_of_contents_doc=}
    {file_path_of_contents_doc_2=}
    {file_path_of_output=}""")

# ソースファイル（JSON形式）を読込
with open(file_path_of_contents_doc, encoding='utf-8') as f:
    contents_doc = json.load(f)


# ドキュメントに対して、影の自動設定の編集を行います
AutoShadowSolver.edit_document(contents_doc)

with open(file_path_of_contents_doc_2, mode='w', encoding='utf-8') as f:
    f.write(json.dumps(contents_doc, indent=4, ensure_ascii=False))

with open(file_path_of_contents_doc_2, mode='r', encoding='utf-8') as f:
    contents_doc = json.load(f)


# ワークブックを新規生成
wb = xl.Workbook()

# ワークシート
ws = wb['Sheet']

# ワークシートへの描画
tr.render_to_worksheet(ws, contents_doc)

# ワークブックの保存            
wb.save(file_path_of_output)
