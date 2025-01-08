"""
白紙の作成
"""

import openpyxl as xl


# ワークブックを新規生成
wb = xl.Workbook()

# ワークブックの保存            
wb.save('./temp/examples/step1_render_empty.xlsx')
