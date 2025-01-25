from ...renderer import fill_rectangle
from ...shared_models import Pillar


def render_all_pillar_rugs(config_doc, contents_doc, ws):
    """全ての柱の敷物の描画
    """

    # 処理しないフラグ
    if 'renderer' in config_doc and (renderer_dict := config_doc['renderer']):
        if 'features' in renderer_dict and (features_dict := renderer_dict['features']):
            if 'pillars' in features_dict and (feature_dict := features_dict['pillars']):
                if 'enabled' in feature_dict:
                    enabled = feature_dict['enabled'] # False 値を取りたい
                    if not enabled:
                        return

    print('🔧　全ての柱の敷物の描画')

    # もし、柱のリストがあれば
    if 'pillars' in contents_doc and (pillars_list := contents_doc['pillars']):

        for pillar_dict in pillars_list:
            pillar_obj = Pillar.from_dict(pillar_dict)

            if 'background' in pillar_dict and (background_dict := pillar_dict['background']):

                if 'varColor' in background_dict and (bg_color := background_dict['varColor']):
                    pillar_bounds_obj = pillar_obj.bounds_obj

                    # 矩形を塗りつぶす
                    fill_rectangle(
                            ws=ws,
                            contents_doc=contents_doc,
                            column_th=pillar_bounds_obj.left_obj.total_of_out_counts_th,
                            row_th=pillar_bounds_obj.top_obj.total_of_out_counts_th,
                            columns=pillar_bounds_obj.width_obj.total_of_out_counts_qty,
                            rows=pillar_bounds_obj.height_obj.total_of_out_counts_qty,
                            color=bg_color)
