from ...share import Card, Pillar, Rectangle, Share, Terminal
from ..translator import Translator


class AutoShadow(Translator):


    def translate_document(self, contents_doc_rw):
        """ドキュメントに対して、影の自動設定の編集を行います

        ['pillars']['cards']['shadow']['varColor] の値が 'auto' なら、
        ['pillars']['cards']['shadow']['varColor] の値を カラーコードに翻訳する
        
        ['pillars']['terminals']['shadow']['varColor] の値が 'auto' なら、
        ['pillars']['terminals']['shadow']['varColor] の値を カラーコードに翻訳する
        
        ['lineTapes']['segments']['shadow']['varColor] の値が 'auto' なら、
        ['lineTapes']['segments']['shadow']['varColor] の値を カラーコードに翻訳する

        Parameters
        ----------
        contents_doc_rw : dict
            読み書き両用
        """

        if 'pillars' in contents_doc_rw and (pillars_list_rw := contents_doc_rw['pillars']):

            for pillar_dict_rw in pillars_list_rw:
                pillar_obj = Pillar.from_dict(pillar_dict_rw)

                if 'cards' in pillar_dict_rw and (card_dict_list_rw := pillar_dict_rw['cards']):

                    for card_dict_rw in card_dict_list_rw:
                        card_obj = Card.from_dict(card_dict_rw)

                        if 'shadow' in card_dict_rw and (shadow_dict_rw := card_dict_rw['shadow']):
                            if 'varColor' in shadow_dict_rw and (shadow_color_value := shadow_dict_rw['varColor']):

                                if shadow_color_value == 'auto':
                                    card_bounds_obj = card_obj.bounds_obj

                                    try:
                                        if solved_var_color_name := AutoShadow._get_auto_shadow(
                                                contents_doc=contents_doc_rw,
                                                column_th=card_bounds_obj.left_obj.total_of_out_counts_th + Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                                row_th=card_bounds_obj.top_obj.total_of_out_counts_th + Share.OUT_COUNTS_THAT_CHANGE_INNING):
                                            shadow_dict_rw['varColor'] = solved_var_color_name
                                    except:
                                        print(f'ERROR: AutoShadow: {card_dict_rw=}')
                                        raise

                # もし、端子のリストがあれば
                if 'terminals' in pillar_dict_rw and (terminals_list_rw := pillar_dict_rw['terminals']):

                    for terminal_dict_rw in terminals_list_rw:
                        terminal_obj = Terminal.from_dict(terminal_dict_rw)
                        terminal_bounds_obj = terminal_obj.bounds_obj

                        if 'shadow' in terminal_dict_rw and (shadow_dict_rw := terminal_dict_rw['shadow']):
                            if 'varColor' in shadow_dict_rw and (shadow_color_value := shadow_dict_rw['varColor']):

                                if shadow_color_value == 'auto':

                                    try:
                                        # 影に自動が設定されていたら、解決する
                                        if solved_var_color_name := AutoShadow._get_auto_shadow(
                                                contents_doc=contents_doc_rw,
                                                column_th=terminal_bounds_obj.left_obj.total_of_out_counts_th + Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                                row_th=terminal_bounds_obj.top_obj.total_of_out_counts_th + Share.OUT_COUNTS_THAT_CHANGE_INNING):
                                            shadow_dict_rw['varColor'] = solved_var_color_name
                                    except:
                                        print(f'ERROR: AutoShadow: {terminal_dict_rw=}')
                                        raise

        # もし、ラインテープのリストがあれば
        if 'lineTapes' in contents_doc_rw and (line_tape_list_rw := contents_doc_rw['lineTapes']):

            for line_tape_dict_rw in line_tape_list_rw:
                # もし、セグメントのリストがあれば
                if 'segments' in line_tape_dict_rw and (segment_list_rw := line_tape_dict_rw['segments']):

                    for segment_dict_rw in segment_list_rw:
                        if 'shadow' in segment_dict_rw and (shadow_dict_rw := segment_dict_rw['shadow']):
                            if 'varColor' in shadow_dict_rw and (shadow_color_value := shadow_dict_rw['varColor']) and shadow_color_value == 'auto':
                                if 'bounds' in segment_dict_rw and (bounds_dict := segment_dict_rw['bounds']):
                                    segment_rect = Rectangle.from_dict(bounds_dict)

                                    # NOTE 影が指定されているということは、浮いているということでもある

                                    try:
                                        # 影に自動が設定されていたら、解決する
                                        if solved_var_color_name := AutoShadow._get_auto_shadow(
                                                contents_doc=contents_doc_rw,
                                                column_th=segment_rect.left_obj.total_of_out_counts_th + Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                                row_th=segment_rect.top_obj.total_of_out_counts_th + Share.OUT_COUNTS_THAT_CHANGE_INNING):
                                            shadow_dict_rw['varColor'] = solved_var_color_name
                                    except:
                                        print(f'ERROR: AutoShadow: {segment_dict_rw=}')
                                        raise


    @staticmethod
    def _get_auto_shadow(contents_doc, column_th, row_th):
        """影に対応する色名を取得"""

        if 'colorSystem' in contents_doc and (color_system_dict := contents_doc['colorSystem']):

            # もし、影の色の対応付けがあれば
            if 'shadowColorMappings' in color_system_dict and (shadow_color_dict := color_system_dict['shadowColorMappings']):

                # もし、柱のリストがあれば
                if 'pillars' in contents_doc and (pillars_list := contents_doc['pillars']):

                    for pillar_dict in pillars_list:
                        pillar_obj = Pillar.from_dict(pillar_dict)

                        # 柱と柱の隙間（隙間柱）は無視する
                        if 'background' not in pillar_dict:
                            continue

                        background_dict = pillar_dict['background']

                        if 'varColor' not in background_dict:
                            continue

                        if not (bg_color := background_dict['varColor']):
                            continue

                        pillar_bounds_obj = pillar_obj.bounds_obj

                        # もし、矩形の中に、指定の点が含まれたなら
                        if pillar_bounds_obj.left_obj.total_of_out_counts_th <= column_th and column_th < pillar_bounds_obj.left_obj.total_of_out_counts_th + pillar_bounds_obj.width_obj.total_of_out_counts_qty and \
                            pillar_bounds_obj.top_obj.total_of_out_counts_th <= row_th and row_th < pillar_bounds_obj.top_obj.total_of_out_counts_th + pillar_bounds_obj.height_obj.total_of_out_counts_qty:

                            # ベースの色に紐づく影の色
                            if bg_color in shadow_color_dict:
                                shadow_color = shadow_color_dict[bg_color]
                                return shadow_color

                            # ベースの色に紐づく影色が見つからない
                            else:
                                return None

        # 該当なし
        if 'paperColor' in shadow_color_dict:
            return shadow_color_dict['paperColor']
        else:
            return None
