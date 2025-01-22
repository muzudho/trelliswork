from ..share import InningsPitched, Pillar, Rectangle, Share


class AutoSplitPillarSolver():


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
                            new_splitting_segments.extend(AutoSplitPillarSolver._split_segment_by_pillar(
                                    contents_doc=contents_doc,
                                    line_tape_segment_list=line_tape_segment_list,
                                    line_tape_segment_dict=line_tape_segment_dict))

        # 削除用ループが終わってから追加する。そうしないと無限ループしてしまう
        for splitting_segments in new_splitting_segments:
            line_tape_segment_list.append(splitting_segments)


    @staticmethod
    def _split_segment_by_pillar(contents_doc, line_tape_segment_list, line_tape_segment_dict):
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
