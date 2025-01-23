from ..share import Card, Pillar, Rectangle, Share, Terminal


class AutoShadowSolver():


    @staticmethod
    def edit_document(contents_doc_rw):
        """ドキュメントに対して、影の自動設定の編集を行います

        ['pillars']['cards']['shadowColor'] の値が 'auto' なら、
        ['pillars']['cards']['shadowColor'] の値を カラーコードに翻訳する
        
        ['pillars']['terminals']['shadowColor'] の値が 'auto' なら、
        ['pillars']['terminals']['shadowColor'] の値を カラーコードに翻訳する
        
        ['lineTapes']['segments']['shadowColor'] の値が 'auto' なら、
        ['lineTapes']['segments']['shadowColor'] の値を カラーコードに翻訳する

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

                        if 'shadowColor' in card_dict_rw and (card_shadow_color := card_dict_rw['shadowColor']):

                            if card_shadow_color == 'auto':
                                card_bounds_obj = card_obj.bounds_obj

                                try:
                                    if solved_var_color_name := AutoShadowSolver._get_auto_shadow(
                                            contents_doc=contents_doc_rw,
                                            column_th=card_bounds_obj.left_obj.total_of_out_counts_th + Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                            row_th=card_bounds_obj.top_obj.total_of_out_counts_th + Share.OUT_COUNTS_THAT_CHANGE_INNING):
                                        card_dict_rw['shadowColor'] = solved_var_color_name
                                except:
                                    print(f'ERROR: edit_document_and_solve_auto_shadow: {card_dict_rw=}')
                                    raise

                # もし、端子のリストがあれば
                if 'terminals' in pillar_dict_rw and (terminals_list := pillar_dict_rw['terminals']):

                    for terminal_dict in terminals_list:
                        terminal_obj = Terminal.from_dict(terminal_dict)
                        terminal_bounds_obj = terminal_obj.bounds_obj

                        if 'shadowColor' in terminal_dict and (terminal_shadow_color := terminal_dict['shadowColor']):

                            if terminal_shadow_color == 'auto':

                                # 影に自動が設定されていたら、解決する
                                if solved_var_color_name := AutoShadowSolver._get_auto_shadow(
                                        contents_doc=contents_doc_rw,
                                        column_th=terminal_bounds_obj.left_obj.total_of_out_counts_th + Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                        row_th=terminal_bounds_obj.top_obj.total_of_out_counts_th + Share.OUT_COUNTS_THAT_CHANGE_INNING):
                                    terminal_dict['shadowColor'] = solved_var_color_name

        # もし、ラインテープのリストがあれば
        if 'lineTapes' in contents_doc_rw and (line_tape_list := contents_doc_rw['lineTapes']):

            for line_tape_dict in line_tape_list:
                # もし、セグメントのリストがあれば
                if 'segments' in line_tape_dict and (segment_list := line_tape_dict['segments']):

                    for segment_dict in segment_list:
                        if 'shadowColor' in segment_dict and (segment_shadow_color := segment_dict['shadowColor']) and segment_shadow_color == 'auto':
                            segment_rect = Rectangle.from_dict(segment_dict)

                            # NOTE 影が指定されているということは、浮いているということでもある

                            # 影に自動が設定されていたら、解決する
                            if solved_var_color_name := AutoShadowSolver._get_auto_shadow(
                                    contents_doc=contents_doc_rw,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th + Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th + Share.OUT_COUNTS_THAT_CHANGE_INNING):
                                segment_dict['shadowColor'] = solved_var_color_name


    @staticmethod
    def _get_auto_shadow(contents_doc, column_th, row_th):
        """影に対応する色名を取得"""

        # もし、影の色の対応付けがあれば
        if 'shadowColorMappings' in contents_doc and (shadow_color_dict := contents_doc['shadowColorMappings']):

            # もし、柱のリストがあれば
            if 'pillars' in contents_doc and (pillars_list := contents_doc['pillars']):

                for pillar_dict in pillars_list:
                    pillar_obj = Pillar.from_dict(pillar_dict)

                    # 柱と柱の隙間（隙間柱）は無視する
                    if 'baseColor' not in pillar_dict or not pillar_dict['baseColor']:
                        continue

                    pillar_bounds_obj = pillar_obj.bounds_obj
                    base_color = pillar_dict['baseColor']

                    # もし、矩形の中に、指定の点が含まれたなら
                    if pillar_bounds_obj.left_obj.total_of_out_counts_th <= column_th and column_th < pillar_bounds_obj.left_obj.total_of_out_counts_th + pillar_bounds_obj.width_obj.total_of_out_counts_qty and \
                        pillar_bounds_obj.top_obj.total_of_out_counts_th <= row_th and row_th < pillar_bounds_obj.top_obj.total_of_out_counts_th + pillar_bounds_obj.height_obj.total_of_out_counts_qty:

                        return shadow_color_dict[base_color]

        # 該当なし
        return shadow_color_dict['paperColor']
