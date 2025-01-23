from ...renderer import fill_rectangle
from ...share import Rectangle


def render_all_rectangles(config_doc, contents_doc, ws):
    """全ての矩形の描画
    """

    # 処理しないフラグ
    if 'renderer' in config_doc and (renderer_dict := config_doc['renderer']):
        if 'features' in renderer_dict and (features_dict := renderer_dict['features']):
            if 'rectangles' in features_dict and (feature_dict := features_dict['rectangles']):
                if 'enabled' in feature_dict:
                    enabled = feature_dict['enabled'] # False 値を取りたい
                    if not enabled:
                        return

    print('🔧　全ての矩形の描画')

    # もし、矩形のリストがあれば
    if 'rectangles' in contents_doc and (rectangles_list := contents_doc['rectangles']):

        for rectangle_dict in rectangles_list:

            bounds_obj = None
            if 'bounds' in rectangle_dict and (bounds_dict := rectangle_dict['bounds']):
                bounds_obj = Rectangle.from_dict(bounds_dict)

                # セル結合
                if 'mergeCells' in rectangle_dict and (is_merge_cells := rectangle_dict['mergeCells']):
                    if is_merge_cells:
                        column_th = bounds_obj.left_obj.total_of_out_counts_th
                        row_th = bounds_obj.top_obj.total_of_out_counts_th
                        columns = bounds_obj.width_obj.total_of_out_counts_qty
                        rows = bounds_obj.height_obj.total_of_out_counts_qty

                        if 0 < columns and 0 < rows and (1 < columns or 1 < rows):
                            column_letter = xl.utils.get_column_letter(column_th)
                            column_letter2 = xl.utils.get_column_letter(column_th + columns - 1)
                            ws.merge_cells(f'{column_letter}{row_th}:{column_letter2}{row_th + rows - 1}')

                if 'color' in rectangle_dict and (bg_color := rectangle_dict['color']):
                    # もし境界線が指定されていれば、描画する
                    if 'xlBorder' in rectangle_dict and (xl_border_dict := rectangle_dict['xlBorder']):
                        draw_xl_border_on_rectangle(
                                ws=ws,
                                xl_border_dict=xl_border_dict,
                                column_th=bounds_obj.left_obj.total_of_out_counts_th,
                                row_th=bounds_obj.top_obj.total_of_out_counts_th,
                                columns=bounds_obj.width_obj.total_of_out_counts_qty,
                                rows=bounds_obj.height_obj.total_of_out_counts_qty)

                    # 矩形を塗りつぶす
                    fill_rectangle(
                            ws=ws,
                            column_th=bounds_obj.left_obj.total_of_out_counts_th,
                            row_th=bounds_obj.top_obj.total_of_out_counts_th,
                            columns=bounds_obj.width_obj.total_of_out_counts_qty,
                            rows=bounds_obj.height_obj.total_of_out_counts_qty,
                            color=bg_color)
