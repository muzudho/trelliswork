import os
import json

from ..shared_models import FilePath
from ..compiler_parts import AutoShadow, AutoSplitSegmentByPillar, Imports, ResolveAliasOfColor, ResolveVarBounds


class Compiler():
    """コンパイラー
    """


    def __init__(self):
        pass


    def compile(self, config, source=None):
        """コンパイル

        staticmethod の方が適切だが
              import trelliswork as tl
              tc = tl.Compiler()
              tc.compile(config="...", source="...")
        のような書き方がしたいのでインスタンスのメソッドとした。

        出力は、オブジェクト（中間）ファイルという形で出力される。
        オブジェクト・ファイルへのパスは、設定ファイルの方に書かれる。

        Parameters
        ----------
        config : str
            設定ファイル（読取専用）へのパス
        source : str
            内容ファイル（読取専用）へのパス
        """

        print(f"🔧　read {config} config file.")
        with open(config, encoding='utf-8') as f:
            config_dict = json.load(f)

        if 'builder' not in config_dict:
            config_dict['builder'] = {}

        if 'compiler' not in config_dict:
            config_dict['compiler'] = {}

        # 引数が指定されていれば、設定ファイルの記述より、引数を優先します
        if source is not None:
            config_dict['compiler']['source'] = source

        print(f"🔧　read {source} source file.")
        with open(source, encoding='utf-8') as f:
            source_dict_rw = json.load(f)

        tc = Compiler()
        tc._compile_by_dict(
                config_dict=config_dict,
                source_dict_rw=source_dict_rw)


    def _compile_by_dict(self, config_dict, source_dict_rw):
        """コンパイル

        Parameters
        ----------
        config_dict : dict
            設定
        source_dict_rw : dict
            読み書き両用
        """

        source_fp = FilePath(config_dict['compiler']['source'])

        if 'compiler' in config_dict and (compiler_dict := config_dict['compiler']):

            def get_object_folder():
                if 'folderForObjects' not in compiler_dict:
                    raise ValueError("""設定ファイルでコンパイラーの処理結果を中間ファイルとして出力する設定にした場合は、['compiler']['folderForObjects']が必要です。""")

                return compiler_dict['folderForObjects']


            if 'prefixForObjectFiles' in compiler_dict and (prefix_for_object_files := compiler_dict['prefixForObjectFiles']) and prefix_for_object_files is not None:
                pass
            else:
                prefix_for_object_files = ''


            if 'parts' in compiler_dict and (parts_dict := compiler_dict['parts']):


                def create_filepath_of_object_file(source_fp, object_file_dict):
                    """中間ファイルのパス作成"""

                    prefix = ''
                    if prefix_for_object_files:
                        prefix = f'{prefix_for_object_files}__'

                    object_suffix = object_file_dict['suffix']
                    basename = f'{prefix}{source_fp.basename_without_ext}__{object_suffix}.json'
                    return os.path.join(get_object_folder(), basename)


                def write_object_file(comment):
                    """中間ファイルの書出し
                    """
                    if 'objectFile' in compiler_part_dict and (object_file_dict := compiler_part_dict['objectFile']):
                        filepath_of_object_file = create_filepath_of_object_file(
                                source_fp=source_fp,
                                object_file_dict=object_file_dict)

                        # ディレクトリーが存在しなければ作成する
                        directory_path = os.path.split(filepath_of_object_file)[0]
                        os.makedirs(directory_path, exist_ok=True)

                        print(f"""\
🔧　write {filepath_of_object_file} object file.
    {comment=}""")

                        with open(filepath_of_object_file, mode='w', encoding='utf-8') as f:
                            f.write(json.dumps(source_dict_rw, indent=4, ensure_ascii=False))


                # ［コンパイラーの部品一覧］
                compiler_part_instance_dict = {
                    'autoSplitSegmentByPillar': AutoSplitSegmentByPillar(),
                    'autoShadow': AutoShadow(),
                    'imports': Imports(),
                    'resolveAliasOfColor': ResolveAliasOfColor(),
                    'resolveVarBounds': ResolveVarBounds(),
                }

                # コンパイラーの部品の実行順序
                if 'orderOfParts' in compiler_dict and (order_of_parts_list := compiler_dict['orderOfParts']):

                    for compiler_part_key in order_of_parts_list:

                        # 各［コンパイラーの部品］
                        #
                        #   ［コンパイラーの部品］は compile_document(source_dict_rw) というインスタンス・メソッドを持つ
                        #
                        compiler_part_dict = parts_dict[compiler_part_key]

                        if compiler_part_key in compiler_part_instance_dict:
                            compiler_part_obj = compiler_part_instance_dict[compiler_part_key]

                            if 'enabled' in compiler_part_dict and (enabled := compiler_part_dict['enabled']) and enabled:
                                # ドキュメントに対して、自動ピラー分割の編集を行います
                                compiler_part_obj.compile_document(
                                        contents_dict_rw=source_dict_rw)

                            # （場合により）中間ファイルの書出し
                            write_object_file(comment=compiler_part_key)
