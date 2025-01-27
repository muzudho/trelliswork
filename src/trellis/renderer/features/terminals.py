from ...renderer import fill_start_terminal, fill_end_terminal
from ...shared_models import Pillar, Terminal


def render_all_terminals(config_doc, contents_doc, ws):
    """全ての端子の描画
    """

    # 処理しないフラグ
    if 'renderer' in config_doc and (renderer_dict := config_doc['renderer']):
        if 'features' in renderer_dict and (features_dict := renderer_dict['features']):
            if 'terminals' in features_dict and (feature_dict := features_dict['terminals']):
                if 'enabled' in feature_dict:
                    enabled = feature_dict['enabled'] # False 値を取りたい
                    if not enabled:
                        return

    print('🔧　全ての端子の描画')

    # もし、柱のリストがあれば
    if 'pillars' in contents_doc and (pillars_list := contents_doc['pillars']):

        for pillar_dict in pillars_list:
            pillar_obj = Pillar.from_dict(pillar_dict)

            # もし、端子のリストがあれば
            if 'terminals' in pillar_dict and (terminals_list := pillar_dict['terminals']):

                for terminal_dict in terminals_list:
                    terminal_obj = Terminal.from_dict(terminal_dict)
                    terminal_bounds_obj = terminal_obj.bounds_obj

                    terminal_pixel_art = terminal_dict['pixelArt']

                    if terminal_pixel_art == 'start':
                        # 始端のドット絵を描く
                        fill_start_terminal(
                            ws=ws,
                            column_th=terminal_bounds_obj.left_th,
                            row_th=terminal_bounds_obj.top_th)

                    elif terminal_pixel_art == 'end':
                        # 終端のドット絵を描く
                        fill_end_terminal(
                            ws=ws,
                            column_th=terminal_bounds_obj.left_th,
                            row_th=terminal_bounds_obj.top_th)
