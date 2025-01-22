from ..share import Card, InningsPitched, Pillar, Rectangle, Share, Terminal


class AutoShadowSolver():

    @staticmethod
    def edit_document(contents_doc):
        """ドキュメントに対して、影の自動設定の編集を行います
        """

        # もし、柱のリストがあれば
        if 'pillars' in contents_doc and (pillars_list := contents_doc['pillars']):

            for pillar_dict in pillars_list:
                pillar_obj = Pillar.from_dict(pillar_dict)

                # もし、カードの辞書があれば
                if 'cards' in pillar_dict and (card_dict_list := pillar_dict['cards']):

                    for card_dict in card_dict_list:
                        card_obj = Card.from_dict(card_dict)

                        if 'shadowColor' in card_dict and (card_shadow_color := card_dict['shadowColor']):

                            if card_shadow_color == 'auto':
                                card_rect_obj = card_obj.rect_obj

                                # 影に自動が設定されていたら、解決する
                                try:
                                    if solved_var_color_name := AutoShadowSolver.resolve_auto_shadow(
                                            contents_doc=contents_doc,
                                            column_th=card_rect_obj.left_obj.total_of_out_counts_th + Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                            row_th=card_rect_obj.top_obj.total_of_out_counts_th + Share.OUT_COUNTS_THAT_CHANGE_INNING):
                                        card_dict['shadowColor'] = solved_var_color_name
                                except:
                                    print(f'ERROR: edit_document_and_solve_auto_shadow: {card_dict=}')
                                    raise

                # もし、端子のリストがあれば
                if 'terminals' in pillar_dict and (terminals_list := pillar_dict['terminals']):

                    for terminal_dict in terminals_list:
                        terminal_obj = Terminal.from_dict(terminal_dict)
                        terminal_rect_obj = terminal_obj.rect_obj

                        if 'shadowColor' in terminal_dict and (terminal_shadow_color := terminal_dict['shadowColor']):

                            if terminal_shadow_color == 'auto':

                                # 影に自動が設定されていたら、解決する
                                if solved_var_color_name := AutoShadowSolver.resolve_auto_shadow(
                                        contents_doc=contents_doc,
                                        column_th=terminal_rect_obj.left_obj.total_of_out_counts_th + Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                        row_th=terminal_rect_obj.top_obj.total_of_out_counts_th + Share.OUT_COUNTS_THAT_CHANGE_INNING):
                                    terminal_dict['shadowColor'] = solved_var_color_name

        # もし、ラインテープのリストがあれば
        if 'lineTapes' in contents_doc and (line_tape_list := contents_doc['lineTapes']):

            for line_tape_dict in line_tape_list:
                # もし、セグメントのリストがあれば
                if 'segments' in line_tape_dict and (segment_list := line_tape_dict['segments']):

                    for segment_dict in segment_list:
                        if 'shadowColor' in segment_dict and (segment_shadow_color := segment_dict['shadowColor']) and segment_shadow_color == 'auto':
                            segment_rect = Rectangle.from_dict(segment_dict)

                            # NOTE 影が指定されているということは、浮いているということでもある

                            # 影に自動が設定されていたら、解決する
                            if solved_var_color_name := AutoShadowSolver.resolve_auto_shadow(
                                    contents_doc=contents_doc,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th + Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th + Share.OUT_COUNTS_THAT_CHANGE_INNING):
                                segment_dict['shadowColor'] = solved_var_color_name


    @staticmethod
    def resolve_auto_shadow(contents_doc, column_th, row_th):
        """影の自動設定を解決する"""

        # もし、影の色の対応付けがあれば
        if 'shadowColorMappings' in contents_doc and (shadow_color_dict := contents_doc['shadowColorMappings']):

            # もし、柱のリストがあれば
            if 'pillars' in contents_doc and (pillars_list := contents_doc['pillars']):

                for pillar_dict in pillars_list:
                    pillar_obj = Pillar.from_dict(pillar_dict)

                    # 柱と柱の隙間（隙間柱）は無視する
                    if 'baseColor' not in pillar_dict or not pillar_dict['baseColor']:
                        continue

                    pillar_rect_obj = pillar_obj.rect_obj
                    base_color = pillar_dict['baseColor']

                    # もし、矩形の中に、指定の点が含まれたなら
                    if pillar_rect_obj.left_obj.total_of_out_counts_th <= column_th and column_th < pillar_rect_obj.left_obj.total_of_out_counts_th + pillar_rect_obj.width_obj.total_of_out_counts_qty and \
                        pillar_rect_obj.top_obj.total_of_out_counts_th <= row_th and row_th < pillar_rect_obj.top_obj.total_of_out_counts_th + pillar_rect_obj.height_obj.total_of_out_counts_qty:

                        return shadow_color_dict[base_color]

        # 該当なし
        return shadow_color_dict['paperColor']


class AutoSplitPillar():


    @staticmethod
    def edit_document(contents_doc):
        """ドキュメントに対して、影の自動設定の編集を行います
        """
        new_splitting_segments = []

        # もし、ラインテープのリストがあれば
        if 'lineTapes' in contents_doc and (line_tape_list := contents_doc['lineTapes']):

            for line_tape_dict in line_tape_list:
                # もし、セグメントのリストがあれば
                if 'segments' in line_tape_dict and (line_tape_segment_list := line_tape_dict['segments']):

                    for line_tape_segment_dict in line_tape_segment_list:
                        # もし、影があれば
                        if 'shadowColor' in line_tape_segment_dict and (shadow_color := line_tape_segment_dict['shadowColor']):
                            # 柱を跨ぐとき、ラインテープを分割します
                            new_splitting_segments.extend(AutoSplitPillar.split_segment_by_pillar(
                                    contents_doc=contents_doc,
                                    line_tape_segment_list=line_tape_segment_list,
                                    line_tape_segment_dict=line_tape_segment_dict))

        # 削除用ループが終わってから追加する。そうしないと無限ループしてしまう
        for splitting_segments in new_splitting_segments:
            line_tape_segment_list.append(splitting_segments)


    @staticmethod
    def split_segment_by_pillar(contents_doc, line_tape_segment_list, line_tape_segment_dict):
        """柱を跨ぐとき、ラインテープを分割します
        NOTE 柱は左から並んでいるものとする
        NOTE 柱の縦幅は十分に広いものとする
        NOTE テープは浮いています
        """

        new_segment_list = []

        #print('🔧　柱を跨ぐとき、ラインテープを分割します')
        segment_rect = Rectangle.from_dict(line_tape_segment_dict)

        direction = line_tape_segment_dict['direction']

        splitting_segments = []


        # 右進でも、左進でも、同じコードでいけるようだ
        if direction in ['after_falling_down.turn_right', 'after_up.turn_right', 'from_here.go_right', 'after_falling_down.turn_left']:

            # もし、柱のリストがあれば
            if 'pillars' in contents_doc and (pillars_list := contents_doc['pillars']):

                # 各柱
                for pillar_dict in pillars_list:
                    pillar_obj = Pillar.from_dict(pillar_dict)
                    pillar_rect_obj = pillar_obj.rect_obj

                    # とりあえず、ラインテープの左端と右端の内側に、柱の右端があるか判定
                    if segment_rect.left_obj.total_of_out_counts_th < pillar_rect_obj.right_obj.total_of_out_counts_th and pillar_rect_obj.right_obj.total_of_out_counts_th < segment_rect.right_obj.total_of_out_counts_th:
                        # 既存のセグメントを削除
                        line_tape_segment_list.remove(line_tape_segment_dict)

                        # 左側のセグメントを新規作成し、新リストに追加
                        # （計算を簡単にするため）width は使わず right を使う
                        left_segment_dict = dict(line_tape_segment_dict)
                        left_segment_dict.pop('width', None)
                        left_segment_dict['right'] = InningsPitched.from_var_value(pillar_rect_obj.right_obj.var_value).offset(-1).var_value
                        new_segment_list.append(left_segment_dict)

                        # 右側のセグメントを新規作成し、既存リストに追加
                        # （計算を簡単にするため）width は使わず right を使う
                        right_segment_dict = dict(line_tape_segment_dict)
                        right_segment_dict.pop('width', None)
                        right_segment_dict['left'] = pillar_rect_obj.right_obj.offset(-1).var_value
                        right_segment_dict['right'] = segment_rect.right_obj.var_value
                        line_tape_segment_list.append(right_segment_dict)
                        line_tape_segment_dict = right_segment_dict          # 入れ替え


        return new_segment_list
