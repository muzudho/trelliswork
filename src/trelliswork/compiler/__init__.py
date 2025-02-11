import os
import json

from ..shared_models import FilePath
from .translators import AutoShadow, AutoSplitSegmentByPillar, Imports, ResolveAliasOfColor, ResolveVarBounds


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
                if 'objectFolder' not in compiler_dict:
                    raise ValueError("""設定ファイルでコンパイラーの処理結果を中間ファイルとして出力する設定にした場合は、['compiler']['objectFolder']が必要です。""")

                return compiler_dict['objectFolder']


            if 'objectFilePrefix' in compiler_dict and (object_file_prefix := compiler_dict['objectFilePrefix']) and object_file_prefix is not None:
                pass
            else:
                object_file_prefix = ''


            if 'tlanslators' in compiler_dict and (translators_dict := compiler_dict['tlanslators']):


                def create_file_path_of_contents_doc_object(source_fp, object_file_dict):
                    """中間ファイルのパス作成"""
                    object_suffix = object_file_dict['suffix']
                    basename = f'{object_file_prefix}__{source_fp.basename_without_ext}__{object_suffix}.json'
                    return os.path.join(get_object_folder(), basename)


                def write_object_file(comment):
                    """中間ファイルの書出し
                    """
                    if 'objectFile' in translator_dict and (object_file_dict := translator_dict['objectFile']):
                        file_path_of_contents_doc_object = create_file_path_of_contents_doc_object(
                                source_fp=source_fp,
                                object_file_dict=object_file_dict)

                        print(f"""\
🔧　write {file_path_of_contents_doc_object} object file.
    {comment=}""")

                        # ディレクトリーが存在しなければ作成する
                        directory_path = os.path.split(file_path_of_contents_doc_object)[0]
                        os.makedirs(directory_path, exist_ok=True)

                        print(f"🔧　write {file_path_of_contents_doc_object} file.")
                        with open(file_path_of_contents_doc_object, mode='w', encoding='utf-8') as f:
                            f.write(json.dumps(source_dict_rw, indent=4, ensure_ascii=False))


                # ［翻訳者一覧］
                translator_object_dict = {
                    'autoSplitSegmentByPillar': AutoSplitSegmentByPillar(),
                    'autoShadow': AutoShadow(),
                    'imports': Imports(),
                    'resolveAliasOfColor': ResolveAliasOfColor(),
                    'resolveVarBounds': ResolveVarBounds(),
                }

                # 翻訳の実行順序
                if 'translationOrder' in compiler_dict and (translation_order_list := compiler_dict['translationOrder']):

                    for translation_key in translation_order_list:

                        # 各［翻訳者］
                        #
                        #   翻訳者は translate_document(source_dict_rw) というインスタンス・メソッドを持つ
                        #
                        translator_dict = translators_dict[translation_key]

                        if translation_key in translator_object_dict:
                            translator_obj = translator_object_dict[translation_key]

                            if 'enabled' in translator_dict and (enabled := translator_dict['enabled']) and enabled:
                                # ドキュメントに対して、自動ピラー分割の編集を行います
                                translator_obj.translate_document(
                                        contents_dict_rw=source_dict_rw)

                            # （場合により）中間ファイルの書出し
                            write_object_file(comment=translation_key)
