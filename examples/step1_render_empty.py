"""
白紙の作成
"""

import openpyxl as xl
import json


# ソースファイル（JSON形式）を読込
json_file_name = './examples/data/new_paper.json'
with open(json_file_name) as f:
    document = json.load(f)


# Trellis では、タテ：ヨコ＝３：３ で、１ユニットセルとします。
# また、上辺、右辺、下辺、左辺に、１セル幅の定規を置きます
length_of_columns = document['canvas']['width'] * 3 + 2
length_of_rows    = document['canvas']['height'] * 3 + 2

print(f"""\
json_file_name = {json_file_name}
canvas
    length_of_columns = {length_of_columns}
    length_of_rows    = {length_of_rows}
""")

# ワークブックを新規生成
wb = xl.Workbook()

# ワークシート
ws = wb['Sheet']

# 行の横幅
for column_th in range(1, length_of_columns + 1):
    column_letter = xl.utils.get_column_letter(column_th)
    print(f"column_letter={column_letter}")
    ws.column_dimensions[column_letter].width = 2.7    # 2.7 characters = about 30 pixels

# 列の高さ
for row_th in range(1, length_of_rows + 1):
    ws.row_dimensions[row_th].height = 15    # 15 points = about 30 pixels


# ワークブックの保存            
wb.save('./temp/examples/step1_new_paper.xlsx')
