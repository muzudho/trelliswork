from ...renderer import fill_rectangle
from ...shared_models import Card, Pillar, Share


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

                    if 'shadow' in card_dict and (shadow_dict := card_dict['shadow']):
                        if 'varColor' in shadow_dict and (shadow_color_value := shadow_dict['varColor']):
                            

                            card_bounds_obj = card_obj.bounds_obj

                            # 端子の影を描く
                            fill_rectangle(
                                    ws=ws,
                                    contents_doc=contents_doc,
                                    column_th=card_bounds_obj.left_th + Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                    row_th=card_bounds_obj.top_th + Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                    columns=card_bounds_obj.width,
                                    rows=card_bounds_obj.height,
                                    color=shadow_color_value)
