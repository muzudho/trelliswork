"""
白紙に柱の頭を追加
"""

import json
import openpyxl as xl
from openpyxl.styles import PatternFill, Font

from src.trellis import trellis_in_src as tr
from src.trellis.compiler import AutoShadowSolver


print('step 4: auto shadow')

# ソースファイル（JSON形式）を読込
json_file_name = './examples/data/battle_sequence_of_unfair_cointoss.step4_auto_shadow.json'
print(f"json_file_name = {json_file_name}")
with open(json_file_name, encoding='utf-8') as f:
    contents_doc = json.load(f)


# ドキュメントに対して、影の自動設定の編集を行います
AutoShadowSolver.edit_document(contents_doc)

json_file_name_2 = './temp/examples/data_step4_battle_sequence_of_unfair_cointoss.step4_auto_shadow.compiled.json'
print(f"write json_file_name_2 = {json_file_name_2}")
with open(json_file_name_2, mode='w', encoding='utf-8') as f:
    f.write(json.dumps(contents_doc, indent=4, ensure_ascii=False))

print(f"read json_file_name_2 = {json_file_name_2}")
with open(json_file_name_2, mode='r', encoding='utf-8') as f:
    contents_doc = json.load(f)


# ワークブックを新規生成
wb = xl.Workbook()

# ワークシート
ws = wb['Sheet']

# ワークシートへの描画
tr.render_to_worksheet(ws, contents_doc)

# ワークブックの保存            
wb.save('./temp/examples/step4_auto_shadow.xlsx')
