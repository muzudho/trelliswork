from ...renderer import fill_rectangle
from ...share import Card, Pillar, Share


def render_shadow_of_all_cards(config_doc, contents_doc, ws):
    """全てのカードの影の描画
    """

    # 処理しないフラグ
    if 'renderer' in config_doc and (renderer_dict := config_doc['renderer']):
        if 'features' in renderer_dict and (features_dict := renderer_dict['features']):
            if 'shadowOfCards' in features_dict and (feature_dict := features_dict['shadowOfCards']):
                if 'enabled' in feature_dict:
                    enabled = feature_dict['enabled'] # False 値を取りたい
                    if not enabled:
                        return

    print('🔧　全てのカードの影の描画')

    # もし、柱のリストがあれば
    if 'pillars' in contents_doc and (pillars_list := contents_doc['pillars']):

        for pillar_dict in pillars_list:
            pillar_obj = Pillar.from_dict(pillar_dict)

            # もし、カードの辞書があれば
            if 'cards' in pillar_dict and (card_dict_list := pillar_dict['cards']):

                for card_dict in card_dict_list:
                    card_obj = Card.from_dict(card_dict)

                    if 'shadowColor' in card_dict:
                        card_shadow_color = card_dict['shadowColor']

                        card_rect_obj = card_obj.rect_obj

                        # 端子の影を描く
                        fill_rectangle(
                                ws=ws,
                                column_th=card_rect_obj.left_obj.total_of_out_counts_th + Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                row_th=card_rect_obj.top_obj.total_of_out_counts_th + Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                columns=card_rect_obj.width_obj.total_of_out_counts_qty,
                                rows=card_rect_obj.height_obj.total_of_out_counts_qty,
                                color=card_shadow_color)
