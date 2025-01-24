import json

from ..translator import Translator


class Imports(Translator):


    def translate_document(self, contents_doc_rw):

        if 'imports' in contents_doc_rw and (imports_list := contents_doc_rw['imports']):

            for import_file_path in imports_list:

                # ファイル（JSON形式）を読込
                print(f"🔧　import {import_file_path} file")
                with open(import_file_path, encoding='utf-8') as f:
                    import_doc = json.load(f)

                if 'exports' in import_doc and (exports_dict := import_doc['exports']):

                    for export_key, export_body in exports_dict.items():
                        print(f'import: {export_key=}')

                        if export_key in contents_doc_rw:
                            # 辞書と辞書をマージにして、新しい辞書とする
                            contents_doc_rw[export_key] = {**contents_doc_rw[export_key], **export_body}
                        else:
                            # 新規追加
                            contents_doc_rw[export_key] = export_body

            # imports の削除
            del contents_doc_rw['imports']
