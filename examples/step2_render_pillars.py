"""
白紙に柱の頭を追加
"""

import json
import openpyxl as xl
from openpyxl.styles import PatternFill, Font
from openpyxl.styles.alignment import Alignment
from src.trellis import trellis_in_src as tr


print('step 2: pillars')

# ソースファイル（JSON形式）を読込
json_file_name = './examples/data/battle_sequence_of_unfair_cointoss.step1_full_manual.json'
print(f"json_file_name = {json_file_name}")
with open(json_file_name, encoding='utf-8') as f:
    document = json.load(f)

# ワークブックを新規生成
wb = xl.Workbook()

# ワークシート
ws = wb['Sheet']

# 全ての矩形の描画
tr.render_all_rectangles(document, ws)

# 全ての柱の敷物の描画
tr.render_all_pillar_rugs(document, ws)

# 全てのカードの影の描画
tr.render_all_card_shadows(document, ws)

# 全ての端子の影の描画
tr.render_all_terminal_shadows(document, ws)

# 全てのカードの描画
tr.render_all_cards(document, ws)

# 全ての端子の描画
tr.render_all_terminals(document, ws)

# 定規の描画
#       柱を上から塗りつぶすように描きます
tr.render_ruler(document, ws)

# ワークブックの保存            
wb.save('./temp/examples/step2_pillars.xlsx')
