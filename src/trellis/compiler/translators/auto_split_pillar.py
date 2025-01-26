import copy

from ...shared_models import InningsPitched, InningsPitched, Pillar, Rectangle, Share
from ..translator import Translator


class AutoSplitSegmentByPillar(Translator):


    def translate_document(self, contents_doc_rw):
        """ドキュメントに対して、影の自動設定の編集を行います

        Parameters
        ----------
        contents_doc_rw : dict
            読み書き両用
        """
        new_splitting_segments = []

        # もし、ラインテープのリストがあれば
        if 'lineTapes' in contents_doc_rw and (line_tape_list_rw := contents_doc_rw['lineTapes']):

            for line_tape_dict_rw in line_tape_list_rw:
                # もし、セグメントのリストがあれば
                if 'segments' in line_tape_dict_rw and (segment_list_rw := line_tape_dict_rw['segments']):

                    for segment_dict_rw in segment_list_rw:
                        # もし、影があれば
                        if 'shadow' in segment_dict_rw and (shadow_dict := segment_dict_rw['shadow']):
                            if 'varColor' in shadow_dict and (shadow_color_value := shadow_dict['varColor']):
                                # 柱を跨ぐとき、ラインテープを分割します
                                new_splitting_segments.extend(
                                        AutoSplitSegmentByPillar._split_segment_by_pillar(
                                                contents_doc=contents_doc_rw,
                                                segment_list_rw=segment_list_rw,
                                                segment_dict_rw=segment_dict_rw))

        # 削除用ループが終わってから追加する。そうしないと無限ループしてしまう
        for splitting_segments in new_splitting_segments:
            segment_list_rw.append(splitting_segments)


    @staticmethod
    def _split_segment_by_pillar(contents_doc, segment_list_rw, segment_dict_rw):
        """柱を跨ぐとき、ラインテープを分割します

        NOTE 柱は左から並んでいるものとする
        NOTE 柱の縦幅は十分に広いものとする
        NOTE テープは浮いています

        Parameters
        ----------
        segment_list_rw : list
            読み書き両用
        """

        new_segment_list_w = []

        #print('🔧　柱を跨ぐとき、ラインテープを分割します')
        segment_rect_obj = None
        if 'varBounds' in segment_dict_rw and (o1_bounds_dict := segment_dict_rw['varBounds']):
            segment_rect_obj = Rectangle.from_var_bounds_dict(o1_bounds_dict)
        elif 'bounds' in segment_dict_rw and (o2_bounds_dict := segment_dict_rw['bounds']):
            segment_rect_obj = Rectangle.from_bounds_dict(o2_bounds_dict)

        if segment_rect_obj:
            direction = segment_dict_rw['direction']

            splitting_segments = []


            # 右進でも、左進でも、同じコードでいけるようだ
            if direction in ['after_falling_down.turn_right', 'after_up.turn_right', 'from_here.go_right', 'after_falling_down.turn_left']:

                if 'pillars' in contents_doc and (pillars_list := contents_doc['pillars']):

                    # 各柱
                    for pillar_dict in pillars_list:
                        pillar_obj = Pillar.from_dict(pillar_dict)
                        pillar_bounds_obj = pillar_obj.bounds_obj

                        # とりあえず、ラインテープの左端と右端の内側に、柱の右端があるか判定
                        if segment_rect_obj.left_obj.total_of_out_counts_th < pillar_bounds_obj.right_obj.total_of_out_counts_th and pillar_bounds_obj.right_obj.total_of_out_counts_th < segment_rect_obj.right_obj.total_of_out_counts_th:
                            # 既存のセグメントを削除
                            segment_list_rw.remove(segment_dict_rw)

                            # 左側のセグメントを新規作成し、新リストに追加
                            # （計算を簡単にするため）width は使わず right を使う
                            o1_segment_dict = copy.deepcopy(segment_dict_rw)
                            o1_bounds_dict = o1_segment_dict['bounds']
                            o1_bounds_dict.pop('width', None)
                            o1_bounds_dict['right'] = pillar_bounds_obj.right_obj.total_of_out_counts_qty - Share.OUT_COUNTS_THAT_CHANGE_INNING
                            new_segment_list_w.append(o1_segment_dict)

                            # 右側のセグメントを新規作成し、既存リストに追加
                            # （計算を簡単にするため）width は使わず right を使う
                            o2_segment_dict = copy.deepcopy(segment_dict_rw)
                            o2_bounds_dict = o2_segment_dict['bounds']
                            o2_bounds_dict.pop('width', None)
                            o2_bounds_dict['left'] = pillar_bounds_obj.right_obj.total_of_out_counts_qty - Share.OUT_COUNTS_THAT_CHANGE_INNING
                            o2_bounds_dict['right'] = segment_rect_obj.right_obj.total_of_out_counts_qty

                            segment_list_rw.append(o2_segment_dict)
                            segment_dict_rw = o2_segment_dict          # 入れ替え


        return new_segment_list_w
