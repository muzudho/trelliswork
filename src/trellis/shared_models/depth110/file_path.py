import os


################
# MARK: FilePath
################
class FilePath():


    def __init__(self, path):
        """ファイルパスを分解します
        """
        self._directory_path = os.path.split(path)[0]
        self._basename_without_ext = os.path.splitext(os.path.basename(path))[0]
        self._extension_with_dot = os.path.splitext(path)[1]

#         print(f"""\
# ★
# {self._directory_path=}
# {self._basename_without_ext=}
# {self._extension_with_dot=}
# """)


    @property
    def directory_path(self):
        return self._directory_path


    @property
    def basename_without_ext(self):
        return self._basename_without_ext


    @property
    def extension_with_dot(self):
        return self._extension_with_dot
