from ...renderer import fill_rectangle
from ...shared_models import VarRectangle, Share


def render_shadow_of_all_line_tapes(config_doc, contents_doc, ws):
    """全てのラインテープの影の描画
    """

    # 処理しないフラグ
    if 'renderer' in config_doc and (renderer_dict := config_doc['renderer']):
        if 'features' in renderer_dict and (features_dict := renderer_dict['features']):
            if 'shadowOfLineTapes' in features_dict and (feature_dict := features_dict['shadowOfLineTapes']):
                if 'enabled' in feature_dict:
                    enabled = feature_dict['enabled'] # False 値を取りたい
                    if not enabled:
                        return

    print('🔧　全てのラインテープの影の描画')

    # もし、ラインテープの配列があれば
    if 'lineTapes' in contents_doc and (line_tape_list := contents_doc['lineTapes']):

        for line_tape_dict in line_tape_list:
            for segment_dict in line_tape_dict['segments']:
                if 'shadow' in segment_dict and (shadow_dict := segment_dict['shadow']):
                    if 'varColor' in shadow_dict and (shadow_color_value := shadow_dict['varColor']):

                        segment_rect = None
                        if 'bounds' in segment_dict and(o2_bounds_dict := segment_dict['bounds']):
                            segment_rect = VarRectangle.from_bounds_dict(o2_bounds_dict)

                        if segment_rect:
                            # 端子の影を描く
                            fill_rectangle(
                                    ws=ws,
                                    contents_doc=contents_doc,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th + Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th + Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                    columns=segment_rect.width_obj.total_of_out_counts_qty,
                                    rows=segment_rect.height_obj.total_of_out_counts_qty,
                                    color=shadow_color_value)
