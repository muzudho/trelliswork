"""
白紙に柱の頭を追加
"""

import openpyxl as xl
from openpyxl.styles import PatternFill, Font
from openpyxl.styles.alignment import Alignment
import json
from src.trellis import trellis_in_src as tr



# ソースファイル（JSON形式）を読込
json_file_name = './examples/data/battle_sequence_of_unfair_cointoss.json'
print(f"json_file_name = {json_file_name}")
with open(json_file_name, encoding='utf-8') as f:
    document = json.load(f)



# ワークブックを新規生成
wb = xl.Workbook()

# ワークシート
ws = wb['Sheet']

# 全ての柱の頭の描画
tr.render_pillar_headers(document, ws)

# 全ての端子の描画
tr.render_terminals(document, ws)

# 定規の描画
#       柱を上から塗りつぶすように描きます
tr.render_ruler(document, ws)

# ワークブックの保存            
wb.save('./temp/examples/step2_pillar_header.xlsx')
