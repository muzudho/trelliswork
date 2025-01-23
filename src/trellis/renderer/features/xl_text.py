from ...share import XlAlignment


def render_all_xl_texts(config_doc, contents_doc, ws):
    """全てのテキストの描画（定規の番号除く）
    """

    # 処理しないフラグ
    #
    #   TODO xlTexts を使わないテキスト出力が短冊にある。仕様を統一したい
    #
    if 'renderer' in config_doc and (renderer_dict := config_doc['renderer']):
        if 'features' in renderer_dict and (features_dict := renderer_dict['features']):
            if 'xlTexts' in features_dict and (feature_dict := features_dict['xlTexts']):
                if 'enabled' in feature_dict:
                    enabled = feature_dict['enabled'] # False 値を取りたい
                    if not enabled:
                        return

    print(f'🔧　全てのテキストの描画')

    # もし、テキストのリストがあれば
    if 'xlTexts' in contents_doc and (xlTexts := contents_doc['xlTexts']):
        for xl_text_dict in xlTexts:

            # テキスト設定
            if 'text' in xl_text_dict and (text := xl_text_dict['text']):

                # 位置
                location_obj = None
                if 'location' in xl_text_dict and (location_dict := xl_text_dict['location']):
                    location_obj = Point.from_dict(location_dict)

                # テキストの位置揃え
                xl_alignment_obj = None
                if 'xlAlignment' in xl_text_dict and (xl_alignment_dict := xl_text_dict['xlAlignment']):
                    xl_alignment_obj = XlAlignment.from_dict(xl_alignment_dict)

                # フォント
                xl_font_obj = None
                if 'xlFont' in xl_text_dict and (xl_font_dict := xl_text_dict['xlFont']):
                    xl_font_obj = XlFont.from_dict(xl_font_dict)

                # テキストを入力する
                print_text(
                        ws=ws,
                        location_obj=location_obj,
                        text=text,
                        xl_alignment_obj=xl_alignment_obj,
                        xl_font_obj=xl_font_obj)
