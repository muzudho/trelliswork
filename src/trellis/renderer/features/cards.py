from ...renderer import draw_rectangle, render_paper_strip
from ...share import Card, Pillar, Share


def render_all_cards(config_doc, contents_doc, ws):
    """全てのカードの描画
    """

    # 処理しないフラグ
    if 'renderer' in config_doc and (renderer_dict := config_doc['renderer']):
        if 'features' in renderer_dict and (features_dict := renderer_dict['features']):
            if 'cards' in features_dict and (feature_dict := features_dict['cards']):
                if 'enabled' in feature_dict:
                    enabled = feature_dict['enabled'] # False 値を取りたい
                    if not enabled:
                        return

    print('🔧　全てのカードの描画')

    # もし、柱のリストがあれば
    if 'pillars' in contents_doc and (pillars_list := contents_doc['pillars']):

        for pillar_dict in pillars_list:
            pillar_obj = Pillar.from_dict(pillar_dict)

            # ［柱の隙間］は無視する
            if 'background' not in pillar_dict:
                continue

            background_dict = pillar_dict['background']

            if 'varColor' not in background_dict:
                continue

            if not background_dict['varColor']:
                continue


            card_list = pillar_dict['cards']

            for card_dict in card_list:
                card_obj = Card.from_dict(card_dict)
                card_bounds_obj = card_obj.bounds_obj

                try:
                    # ヘッダーの矩形の枠線を描きます
                    draw_rectangle(
                            ws=ws,
                            column_th=card_bounds_obj.left_obj.total_of_out_counts_th,
                            row_th=card_bounds_obj.top_obj.total_of_out_counts_th,
                            columns=card_bounds_obj.width_obj.total_of_out_counts_qty,
                            rows=card_bounds_obj.height_obj.total_of_out_counts_qty)
                except:
                    print(f'ERROR: render_all_cards: {card_dict=}')
                    raise

                if 'paperStrips' in card_dict:
                    paper_strip_list = card_dict['paperStrips']

                    for index, paper_strip in enumerate(paper_strip_list):

                        # 短冊１行の描画
                        render_paper_strip(
                                ws=ws,
                                contents_doc=contents_doc,
                                paper_strip=paper_strip,
                                column_th=card_bounds_obj.left_obj.total_of_out_counts_th,
                                row_th=index * Share.OUT_COUNTS_THAT_CHANGE_INNING + card_bounds_obj.top_obj.total_of_out_counts_th,
                                columns=card_bounds_obj.width_obj.total_of_out_counts_qty,
                                rows=card_bounds_obj.height_obj.total_of_out_counts_qty)
