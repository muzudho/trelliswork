class XlAlignment():
    """Excel 用テキストの位置揃え
    """


    @staticmethod
    def from_dict(xl_alignment_dict):
        """辞書を元に生成

        📖 [openpyxl.styles.alignment module](https://openpyxl.readthedocs.io/en/latest/api/openpyxl.styles.alignment.html)
        horizontal: Value must be one of {‘fill’, ‘left’, ‘distributed’, ‘justify’, ‘center’, ‘general’, ‘centerContinuous’, ‘right’}
        vertical: Value must be one of {‘distributed’, ‘justify’, ‘center’, ‘bottom’, ‘top’}
        """
        xlHorizontal = None
        xlVertical = None
        if 'xlHorizontal' in xl_alignment_dict:
            xlHorizontal = xl_alignment_dict['xlHorizontal']

        if 'xlVertical' in xl_alignment_dict:
            xlVertical = xl_alignment_dict['xlVertical']

        return XlAlignment(
                xlHorizontal=xlHorizontal,
                xlVertical=xlVertical)


    def __init__(self, xlHorizontal, xlVertical):
        self._xl_horizontal = xlHorizontal
        self._xl_vertical = xlVertical


    @property
    def xlHorizontal(self):
        return self._xl_horizontal


    @property
    def xlVertical(self):
        return self._xl_vertical
