from ...renderer import fill_rectangle
from ...share import ColorSystem, Rectangle, Share


def render_all_line_tapes(config_doc, contents_doc, ws):
    """全てのラインテープの描画
    """

    # 処理しないフラグ
    if 'renderer' in config_doc and (renderer_dict := config_doc['renderer']):
        if 'features' in renderer_dict and (features_dict := renderer_dict['features']):
            if 'lineTapes' in features_dict and (feature_dict := features_dict['lineTapes']):
                if 'enabled' in feature_dict:
                    enabled = feature_dict['enabled'] # False 値を取りたい
                    if not enabled:
                        return

    print('🔧　全てのラインテープの描画')

    # もし、ラインテープの配列があれば
    if 'lineTapes' in contents_doc and (line_tape_list := contents_doc['lineTapes']):

        # 各ラインテープ
        for line_tape_dict in line_tape_list:

            line_tape_outline_color = None
            if 'outlineColor' in line_tape_dict:
                line_tape_outline_color = line_tape_dict['outlineColor']

            # 各セグメント
            for segment_dict in line_tape_dict['segments']:

                line_tape_direction = None
                if 'direction' in segment_dict:
                    line_tape_direction = segment_dict['direction']

                if 'color' in segment_dict:
                    line_tape_color = segment_dict['color']

                    segment_rect = Rectangle.from_dict(segment_dict)

                    # ラインテープを描く
                    fill_rectangle(
                            ws=ws,
                            contents_doc=contents_doc,
                            column_th=segment_rect.left_obj.total_of_out_counts_th,
                            row_th=segment_rect.top_obj.total_of_out_counts_th,
                            columns=segment_rect.width_obj.total_of_out_counts_qty,
                            rows=segment_rect.height_obj.total_of_out_counts_qty,
                            color=line_tape_color)

                    # （あれば）アウトラインを描く
                    if line_tape_outline_color and line_tape_direction:
                        outline_fill_obj = ColorSystem.var_color_name_to_fill_obj(
                                contents_doc=contents_doc,
                                var_color_name=line_tape_outline_color)

                        # （共通処理）垂直方向
                        if line_tape_direction in ['from_here.falling_down', 'after_go_right.turn_falling_down', 'after_go_left.turn_up', 'after_go_left.turn_falling_down']:
                            # 左辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    contents_doc=contents_doc,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th - 1,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th + 1,
                                    columns=1,
                                    rows=segment_rect.height_obj.total_of_out_counts_qty - 2,
                                    color=line_tape_outline_color)

                            # 右辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    contents_doc=contents_doc,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th + segment_rect.width_obj.total_of_out_counts_qty,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th + 1,
                                    columns=1,
                                    rows=segment_rect.height_obj.total_of_out_counts_qty - 2,
                                    color=line_tape_outline_color)

                        # （共通処理）水平方向
                        elif line_tape_direction in ['after_falling_down.turn_right', 'continue.go_right', 'after_falling_down.turn_left', 'continue.go_left', 'after_up.turn_right', 'from_here.go_right']:
                            # 上辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    contents_doc=contents_doc,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th + Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th - 1,
                                    columns=segment_rect.width_obj.total_of_out_counts_qty - 2 * Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                    rows=1,
                                    color=line_tape_outline_color)

                            # 下辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    contents_doc=contents_doc,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th + Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th + segment_rect.height_obj.total_of_out_counts_qty,
                                    columns=segment_rect.width_obj.total_of_out_counts_qty - 2 * Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                    rows=1,
                                    color=line_tape_outline_color)

                        # ここから落ちていく
                        if line_tape_direction == 'from_here.falling_down':
                            # 左辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    contents_doc=contents_doc,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th - 1,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th,
                                    columns=1,
                                    rows=1,
                                    color=line_tape_outline_color)

                            # 右辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    contents_doc=contents_doc,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th + segment_rect.width_obj.total_of_out_counts_qty,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th,
                                    columns=1,
                                    rows=1,
                                    color=line_tape_outline_color)

                        # 落ちたあと、右折
                        elif line_tape_direction == 'after_falling_down.turn_right':
                            # 左辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    contents_doc=contents_doc,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th - 1,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th - 1,
                                    columns=1,
                                    rows=2,
                                    color=line_tape_outline_color)

                            # 下辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    contents_doc=contents_doc,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th - 1,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th + 1,
                                    columns=Share.OUT_COUNTS_THAT_CHANGE_INNING + 1,
                                    rows=1,
                                    color=line_tape_outline_color)

                        # そのまま右進
                        elif line_tape_direction == 'continue.go_right':
                            # 上辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    contents_doc=contents_doc,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th - Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th - 1,
                                    columns=2 * Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                    rows=1,
                                    color=line_tape_outline_color)

                            # 下辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    contents_doc=contents_doc,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th - Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th + 1,
                                    columns=2 * Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                    rows=1,
                                    color=line_tape_outline_color)

                        # 右進から落ちていく
                        elif line_tape_direction == 'after_go_right.turn_falling_down':
                            # 上辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    contents_doc=contents_doc,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th - Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th - 1,
                                    columns=2 * Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                    rows=1,
                                    color=line_tape_outline_color)

                            # 下辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    contents_doc=contents_doc,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th - Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th + 1,
                                    columns=Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                    rows=1,
                                    color=line_tape_outline_color)

                            # 右辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    contents_doc=contents_doc,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th + segment_rect.width_obj.total_of_out_counts_qty,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th - 1,
                                    columns=1,
                                    rows=2,
                                    color=line_tape_outline_color)

                        # 落ちたあと左折
                        elif line_tape_direction == 'after_falling_down.turn_left':
                            # 右辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    contents_doc=contents_doc,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th + segment_rect.width_obj.total_of_out_counts_qty,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th - 1,
                                    columns=1,
                                    rows=2,
                                    color=line_tape_outline_color)

                            # 下辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    contents_doc=contents_doc,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th + segment_rect.width_obj.total_of_out_counts_qty - Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th + 1,
                                    columns=Share.OUT_COUNTS_THAT_CHANGE_INNING + 1,
                                    rows=1,
                                    color=line_tape_outline_color)

                        # そのまま左進
                        elif line_tape_direction == 'continue.go_left':
                            # 上辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    contents_doc=contents_doc,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th + Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th - 1,
                                    columns=segment_rect.width_obj.total_of_out_counts_qty,
                                    rows=1,
                                    color=line_tape_outline_color)

                            # 下辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    contents_doc=contents_doc,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th + Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th + 1,
                                    columns=segment_rect.width_obj.total_of_out_counts_qty,
                                    rows=1,
                                    color=line_tape_outline_color)

                        # 左進から上っていく
                        elif line_tape_direction == 'after_go_left.turn_up':
                            # 下辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    contents_doc=contents_doc,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th + segment_rect.height_obj.total_of_out_counts_qty,
                                    columns=2 * Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                    rows=1,
                                    color=line_tape_outline_color)

                            # 左辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    contents_doc=contents_doc,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th - 1,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th + segment_rect.height_obj.total_of_out_counts_qty - 2,
                                    columns=1,
                                    rows=3,
                                    color=line_tape_outline_color)

                            # 右辺（横長）を描く
                            fill_rectangle(
                                    ws=ws,
                                    contents_doc=contents_doc,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th + Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th + segment_rect.height_obj.total_of_out_counts_qty - 2,
                                    columns=Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                    rows=1,
                                    color=line_tape_outline_color)

                        # 上がってきて右折
                        elif line_tape_direction == 'after_up.turn_right':
                            # 左辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    contents_doc=contents_doc,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th - 1,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th,
                                    columns=1,
                                    rows=1,
                                    color=line_tape_outline_color)

                            # 上辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    contents_doc=contents_doc,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th - 1,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th - 1,
                                    columns=Share.OUT_COUNTS_THAT_CHANGE_INNING + 1,
                                    rows=1,
                                    color=line_tape_outline_color)

                        # 左進から落ちていく
                        elif line_tape_direction == 'after_go_left.turn_falling_down':
                            # 上辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    contents_doc=contents_doc,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th - 1,
                                    columns=2 * Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                    rows=1,
                                    color=line_tape_outline_color)

                            # 左辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    contents_doc=contents_doc,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th - 1,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th - 1,
                                    columns=1,
                                    rows=segment_rect.height_obj.total_of_out_counts_qty,
                                    color=line_tape_outline_color)

                            # 右辺（横長）を描く
                            fill_rectangle(
                                    ws=ws,
                                    contents_doc=contents_doc,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th + Share.OUT_COUNTS_THAT_CHANGE_INNING + 1,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th + 1,
                                    columns=Share.OUT_COUNTS_THAT_CHANGE_INNING - 1,
                                    rows=1,
                                    color=line_tape_outline_color)

                        # ここから右進
                        elif line_tape_direction == 'from_here.go_right':
                            # 上辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    contents_doc=contents_doc,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th - 1,
                                    columns=Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                    rows=1,
                                    color=line_tape_outline_color)

                            # 下辺を描く
                            fill_rectangle(
                                    ws=ws,
                                    contents_doc=contents_doc,
                                    column_th=segment_rect.left_obj.total_of_out_counts_th,
                                    row_th=segment_rect.top_obj.total_of_out_counts_th + 1,
                                    columns=Share.OUT_COUNTS_THAT_CHANGE_INNING,
                                    rows=1,
                                    color=line_tape_outline_color)
