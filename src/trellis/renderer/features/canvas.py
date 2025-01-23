import openpyxl as xl

from ...share import Canvas


def render_canvas(config_doc, contents_doc, ws):
    """キャンバスの編集
    """

    # 処理しないフラグ
    if 'renderer' in config_doc and (renderer_dict := config_doc['renderer']):
        if 'features' in renderer_dict and (features_dict := renderer_dict['features']):
            if 'canvas' in features_dict and (feature_dict := features_dict['canvas']):
                if 'enabled' in feature_dict:
                    enabled = feature_dict['enabled'] # False 値を取りたい
                    if not enabled:
                        return

    print("🔧　キャンバスの編集")

    # ウィンドウ枠の固定
    ws.freeze_panes = 'C2'

    # Trellis では、タテ：ヨコ＝３：３ で、１ユニットセルとします。
    # また、上辺、右辺、下辺、左辺に、１セル幅の定規を置きます
    canvas_obj = Canvas.from_dict(contents_doc['canvas'])
    canvas_rect = canvas_obj.bounds_obj

    # 横幅または縦幅が１アウト未満の場合は、定規は描画しません
    if canvas_rect.width_obj.total_of_out_counts_qty < 1 or canvas_rect.height_obj.total_of_out_counts_qty < 1:
        return

    # 行の横幅
    for column_th in range(
            canvas_rect.left_obj.total_of_out_counts_th,
            canvas_rect.left_obj.total_of_out_counts_th + canvas_rect.width_obj.total_of_out_counts_qty):
        column_letter = xl.utils.get_column_letter(column_th)
        ws.column_dimensions[column_letter].width = 2.7    # 2.7 characters = about 30 pixels

    # 列の高さ
    for row_th in range(
            canvas_rect.top_obj.total_of_out_counts_th,
            canvas_rect.top_obj.total_of_out_counts_th + canvas_rect.height_obj.total_of_out_counts_qty):
        ws.row_dimensions[row_th].height = 15    # 15 points = about 30 pixels
