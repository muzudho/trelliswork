from .share import Share


class InningsPitched():
    """投球回。
    トレリスでは、セル番号を指定するのに使っている
    """


    @staticmethod
    def from_integer_and_decimal_part(integer_part, decimal_part):
        """整数部と小数部を指定
        """
        return InningsPitched(integer_part=integer_part, decimal_part=decimal_part)


    @staticmethod
    def from_var_value(var_value):

        try:
            # "100" が来たら 100 にする
            var_value = int(var_value)
        except ValueError:
            pass

        if isinstance(var_value, int):
            return InningsPitched(
                    integer_part=var_value,
                    decimal_part=0)

        elif isinstance(var_value, str):
            integer_part, decimal_part = map(int, var_value.split('o', 2))
            return InningsPitched(
                    integer_part=integer_part,
                    decimal_part=decimal_part)

        else:
            raise ValueError(f'{type(var_value)=} {var_value=}')

        return InningsPitched(var_value)


    def __init__(self, integer_part, decimal_part):
        self._integer_part = integer_part
        self._decimal_part = decimal_part

        if self._decimal_part == 0:
            self._var_value = self._integer_part
        else:
            self._var_value = f'{self._integer_part}o{self._decimal_part}'

        self._total_of_out_counts_qty = self._integer_part * Share.OUT_COUNTS_THAT_CHANGE_INNING + self._decimal_part


    @property
    def var_value(self):
        """投球回の整数だったり、"3o2" 形式の文字列だったりします
        """
        return self._var_value


    @property
    def integer_part(self):
        """投球回の整数部"""
        return self._integer_part


    @property
    def decimal_part(self):
        """投球回の小数部"""
        return self._decimal_part


    @property
    def total_of_out_counts_qty(self):
        """0から始まるアウト・カウントの総数
        """
        return self._total_of_out_counts_qty


    @property
    def total_of_out_counts_th(self):
        """1から始まるアウト・カウントの総数
        """
        return self._total_of_out_counts_qty + 1


    def offset(self, var_value):
        """この投球回に、引数を加算した数を算出して返します"""
        l = self                       # Left operand
        r = InningsPitched.from_var_value(var_value)  # Right operand
        sum_decimal_part = l.decimal_part + r.decimal_part
        integer_part = l.integer_part + r.integer_part + sum_decimal_part // Share.OUT_COUNTS_THAT_CHANGE_INNING
        return InningsPitched.from_integer_and_decimal_part(
                integer_part=integer_part,
                decimal_part=sum_decimal_part % Share.OUT_COUNTS_THAT_CHANGE_INNING)
