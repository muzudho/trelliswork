import json

from ..compiler.part import Part


class Imports(Part):


    def compile_document(self, contents_dict_rw):

        if 'imports' in contents_dict_rw and (imports_list := contents_dict_rw['imports']):

            for import_file_path in imports_list:

                # ファイル（JSON形式）を読込
                print(f"🔧　import {import_file_path} file.")
                with open(import_file_path, encoding='utf-8') as f:
                    import_doc = json.load(f)

                if 'exports' in import_doc and (exports_dict := import_doc['exports']):

                    for export_key, export_body in exports_dict.items():
                        #print(f'import: {export_key=}')

                        if export_key in contents_dict_rw:
                            # 辞書と辞書をマージにして、新しい辞書とする。重複した場合は、後ろに指定した辞書の方で上書きする
                            contents_dict_rw[export_key] = {**export_body, **contents_dict_rw[export_key]}
                        else:
                            # 新規追加
                            contents_dict_rw[export_key] = export_body

            # imports の削除
            del contents_dict_rw['imports']
