from ...renderer import fill_rectangle

from ...shared_models.pillar import Pillar
from ...shared_models.share import Share
from ...shared_models.terminal import Terminal


def render_shadow_of_all_terminals(config_doc, contents_doc, ws):
    """全ての端子の影の描画
    """

    # 処理しないフラグ
    if 'renderer' in config_doc and (renderer_dict := config_doc['renderer']):
        if 'features' in renderer_dict and (features_dict := renderer_dict['features']):
            if 'shadowOfTerminals' in features_dict and (feature_dict := features_dict['shadowOfTerminals']):
                if 'enabled' in feature_dict:
                    enabled = feature_dict['enabled'] # False 値を取りたい
                    if not enabled:
                        return

    print('🔧　全ての端子の影の描画')

    # もし、柱のリストがあれば
    if 'pillars' in contents_doc and (pillars_list := contents_doc['pillars']):

        for pillar_dict in pillars_list:
            pillar_obj = Pillar.from_dict(pillar_dict)

            # もし、端子のリストがあれば
            if 'terminals' in pillar_dict and (terminals_list := pillar_dict['terminals']):

                for terminal_dict in terminals_list:
                    terminal_obj = Terminal.from_dict(terminal_dict)
                    terminal_bounds_obj = terminal_obj.bounds_obj

                    if 'shadow' in terminal_dict and (shadow_dict := terminal_dict['shadow']):
                        if 'varColor' in shadow_dict and (shadow_color_value := shadow_dict['varColor']):

                            # 端子の影を描く
                            fill_rectangle(
                                    ws=ws,
                                    contents_doc=contents_doc,
                                    column_th=terminal_bounds_obj.left_obj.total_of_out_counts_th + Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                    row_th=terminal_bounds_obj.top_obj.total_of_out_counts_th + Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                    columns=9,
                                    rows=9,
                                    color=shadow_color_value)
