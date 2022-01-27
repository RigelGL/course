import texttable
from typing import *


class Value:
    """
    Значение, которое имеет переменные и/или постоянные затрты
    Может иметь потомков
    """
    __slots__ = ['name', '_const', '_variable', 'children', '_display_name']

    def __init__(self, name: str, const: float = 0, variable: float = 0, display_name: str = None):
        self.name = name
        self._const = const
        self._variable = variable
        self.children = None
        self._display_name = display_name

    def add_child(self, variable):
        """
        Добавление потомка
        :param variable: Value тип данных
        :return: Value, переданное в функцию
        """
        if self.children is None:
            self.children = []
        self.children.append(variable)
        return variable

    @property
    def const(self) -> float:
        """
        Возвращаяет постоянные затраты
        """
        if self.children is None:
            return self._const
        s = 0
        for child in self.children:
            s += child.const
        return s

    @property
    def variable(self) -> float:
        """
        Возвращаяет переменные затраты
        """
        if self.children is None:
            return self._variable
        s = 0
        for child in self.children:
            s += child.variable
        return s

    @property
    def total(self) -> float:
        """
        Возвращаяет полные затраты
        """
        return self.const + self.variable

    def head(self, deep: int = 0):
        if deep == 0 or self.children is None:
            return Value(self.name, self.const, self.variable, self._display_name)
        else:
            nc = Value(self.name, display_name=self._display_name)
            for e in self.children:
                nc.add_child(e.head(deep - 1))
            return nc

    def __getitem__(self, item):
        if type(item) is not str:
            return None
        if item == self.name:
            return self
        if self.children is None:
            return None
        for e in self.children:
            if e.name == item:
                return e
        for e in self.children:
            r = e[item]
            if r:
                return r
        return None

    def __str__(self, deep=0):
        c = self.const
        v = self.variable
        if c > 0 and v > 0:
            s = '  ' * deep + '{}, постоянные: {:,.2f}, переменные: {:,.2f}, все: {:,.2f}'.format(self._display_name if self._display_name else self.name, c, v, c + v)
        elif c == 0:
            s = '  ' * deep + '{}, переменные: {:,.2f}'.format(self._display_name if self._display_name else self.name, v)
        else:
            s = '  ' * deep + '{}, постоянные: {:,.2f}'.format(self._display_name if self._display_name else self.name, c)
        if self.children is not None:
            for child in self.children:
                s += '\n' + child.__str__(deep + 1)
        return s

    def __add__(self, other):
        if type(other) == Value:
            return self.total + other.total

        return self.total + other


class _TableRow:
    __slots__ = ['table', 'data']

    def __init__(self, table, data):
        self.table = table
        self.data = data

    def __len__(self):
        return len(self.data)

    def __getitem__(self, item):
        if type(item) == int:
            if item < 0 or item >= len(self.data):
                print('error, index {} out of range [0, {}]'.format(item, len(self.data) - 1))
                return 0
            return self.data[item]
        elif type(item) == str:
            if item not in self.table.headers_map:
                print('error, name {} is not exists in table. Header is: {}'.format(item, self.table.headers))
                return 0
            return self.data[self.table.headers_map[item]]


class Table:
    """
    Таблица, можно извлекать столбцы по имени,
    и считать сумму всех строк с помощью callback
    """
    __slots__ = ['headers', 'headers_map', 'rows']

    def __init__(self, *args):
        self.headers: List[str] = []
        self.headers_map: dict = {}
        total = 0
        for i in args:
            if type(i) != str:
                print('Ignored headers', i)
                continue
            self.headers.append(i)
            self.headers_map[i] = total
            total += 1
        self.rows: List[_TableRow] = []

    def add_row(self, *args):
        """
        Добавление строки в таблицу

        :param args: значения в таблице
        """
        if len(args) != len(self.headers):
            print('error, row len != headers len', len(args), len(self.headers))
            return
        self.rows.append(_TableRow(self, args))

    def get_column(self, name):
        """
        Возвращает столбец по имени

        :param name: названеи столбца
        :return: массив значений столбца, или пустой массив в случае ошибки
        """
        if name not in self.headers_map.keys():
            print('header not found')
            return []
        target = self.headers_map[name]
        return [e[target] for e in self.rows]

    def calculate_sum(self, callback):
        """
        Вычисляет сумму, применяя переданную функцию к каждой строке.
        Функции передаётся строка в качестве аргумента

        :param callback: функция для вычисления значения строки
        :return: сумма всех строк
        """
        s = 0
        for i in self.rows:
            s += callback(i)
        return s

    def find(self, param_name, value) -> _TableRow:
        for row in self.rows:
            if row[param_name] == value:
                return row

    def __len__(self):
        return len(self.rows)

    def __str__(self):
        t = texttable.Texttable(max_width=200)
        t.header(self.headers)
        for row in self.rows:
            t.add_row(row.data)
        return t.draw()


class PercentTable:
    """
    Таблица, считающая процентное соотношение для каждого элемента
    """
    __slots__ = ['rows']

    def __init__(self):
        self.rows: List[Value] = []

    def add_row(self, variable):
        self.rows.append(variable)

    def _get_percent_text_for_value(self, value: [Value, float, int], total: float, deep: int = 0):
        if type(value) is not Value:
            return '  ' * deep + '{:.2f}%'.format(value / total * 100)

        local = '  ' * deep + '{:.2f}%'.format(value.total / total * 100)
        if value.children is not None:
            for sub in value.children:
                local += '\n' + self._get_percent_text_for_value(sub, total, deep + 1)
        return local

    def __str__(self):
        total = 0
        for i in self.rows:
            total += i.total

        t = texttable.Texttable(max_width=200)
        t.header(['value', 'percent'])

        for i in self.rows:
            t.add_row([i, self._get_percent_text_for_value(i, total)])

        return t.draw()


class _PerPercentTableRow:
    __slots__ = ['name', 'percent', 'data', '_per_percent_table']

    def __init__(self, name, percent, per_percent_table, data=None, ):
        self.name = name
        self.percent = percent
        self.data = data
        self._per_percent_table = per_percent_table

    @property
    def amount(self):
        return self._per_percent_table._get_amount(self.percent)


class PerPercentTable:
    """
    Таблица, считающая значения по проценту
    """
    __slots__ = ['rows', '_initial_value', '_minimum_is_one', '_total_percent', '_normalize', '_rows_map']

    def __init__(self, initial_value: [int, float], minimum_is_one: bool = False, normalize: bool = False):
        """

        :param initial_value: начальное значение, от которого отсчитываются остальные
        :param minimum_is_one: True, если минимальное значение должно быть не ниже 1
        :param normalize: нужно ли нормализовать проценты строк
        """
        self.rows: List[_PerPercentTableRow] = []
        self._initial_value = initial_value
        self._minimum_is_one = minimum_is_one
        self._total_percent = 0.0
        self._normalize = normalize
        self._rows_map: dict = {}

    def add_row(self, name: str, percent: float, data=None):
        if name in self._rows_map.keys():
            print('error, double value', name)
            return
        r = _PerPercentTableRow(name, percent, self, data)
        self.rows.append(r)
        self._rows_map[name] = r
        self._total_percent += percent

    @property
    def total(self):
        """
        Возвращает сумму всех строк
        """
        total = 0
        for i in self.rows:
            np = i.percent / self._total_percent if self._normalize else i.percent
            v = np * self._initial_value
            if type(self._initial_value) == int:
                v = max(1 if self._minimum_is_one else 0, round(v))
            total += v
        return total

    def _get_amount(self, percent):
        np = percent / self._total_percent if self._normalize else percent
        v = np * self._initial_value
        if type(self._initial_value) == int:
            v = max(1 if self._minimum_is_one else 0, round(v))
        return v

    def calc_sum(self, callback):
        s = 0
        for row in self.rows:
            np = row.percent / self._total_percent if self._normalize else row.percent
            v = np * self._initial_value
            if type(self._initial_value) == int:
                v = max(1 if self._minimum_is_one else 0, round(v))
            s += callback(v, row.data)
        return s

    def clone(self, initial_value: [int, float]):
        r = PerPercentTable(initial_value, self._minimum_is_one, self._normalize)
        for row in self.rows:
            r.add_row(row.name, row.percent, row.data)
        return r

    def __len__(self):
        return len(self.rows)

    def __str__(self):
        t = texttable.Texttable(max_width=200)
        t.set_cols_dtype([str] * 4)
        t.header(['name', 'percent', 'value', 'data'])
        total = 0

        for i in self.rows:
            np = i.percent / self._total_percent if self._normalize else i.percent
            v = np * self._initial_value
            if type(self._initial_value) == int:
                v = max(1 if self._minimum_is_one else 0, round(v))
            total += v
            t.add_row([i.name, '{:.2f}'.format(np * 100), ('{:,.2f}' if type(self._initial_value) == float else '{:,}').format(v), str(i.data)])
        t.add_row(['total', 100.0 if self._normalize else self._total_percent,
                   ('{:,.2f}' if type(self._initial_value) == float else '{:,}').format(total), ''])
        return t.draw()


class CalculateTable:
    """
    Вычисляет значения результата по массиву переданных аргументов и callback функции
    """
    __slots__ = ['input_data', 'output_data', 'callback', 'context']

    def __init__(self, input_data: List, callback, context=None):
        self.input_data = input_data
        self.callback = callback
        self.context = context
        if context is None:
            self.output_data = [callback(e) for e in input_data]
        else:
            self.output_data = [callback(context, e) for e in input_data]


    @property
    def items(self):
        return zip(self.input_data, self.output_data)

    def __str__(self):
        t = texttable.Texttable(max_width=200)
        t.header([str(e) for e in self.input_data])
        t.add_row([str(e) for e in self.output_data])
        return t.draw()


class ActivePassive:
    __slots__ = [
        'NMA', 'OS',
        'K_ob_sr_pr_zap', 'K_ob_nez_pr', 'K_ob_got_prod', 'K_ob_RBP', 'K_ob_extra', 'debitor_dolg', 'K_ob_ds',
        'ustavnoy_kapital', 'dobavochniy_kapital', 'reservniy_kapital', 'neraspred_pribil',
        'doldosroch_zaemn_sredstva',
        'kratkosroch_zaem_sredstva', 'kratkosroch_prochee'
    ]

    def __init__(self):
        self.NMA = 0
        self.OS = 0
        self.K_ob_sr_pr_zap = 0
        self.K_ob_nez_pr = 0
        self.K_ob_got_prod = 0
        self.K_ob_RBP = 0
        self.K_ob_extra = 0
        self.debitor_dolg = 0
        self.K_ob_ds = 0
        self.ustavnoy_kapital = 0
        self.dobavochniy_kapital = 0
        self.reservniy_kapital = 0
        self.neraspred_pribil = 0
        self.doldosroch_zaemn_sredstva = 0
        self.kratkosroch_zaem_sredstva = 0
        self.kratkosroch_prochee = 0

    @property
    def r1(self):
        return self.NMA + self.OS

    @property
    def r_K_ob_zap(self):
        return self.K_ob_sr_pr_zap + self.K_ob_nez_pr + self.K_ob_got_prod + self.K_ob_RBP + self.K_ob_extra

    @property
    def r2(self):
        return self.r_K_ob_zap + self.debitor_dolg + self.K_ob_ds

    @property
    def r3(self):
        return self.ustavnoy_kapital + self.dobavochniy_kapital + self.reservniy_kapital + self.neraspred_pribil

    @property
    def r4(self):
        return self.doldosroch_zaemn_sredstva

    @property
    def r5(self):
        return self.kratkosroch_zaem_sredstva + self.kratkosroch_prochee

    @property
    def active(self):
        return self.r1 + self.r2

    @property
    def passive(self):
        return self.r3 + self.r4 + self.r5

    def to_table(self):
        return [
            ['1. Внеоборотные активы', None, '3. Капитал и резервы', None],
            ['Нематериальные активы', self.NMA, 'Уставный капитал', self.ustavnoy_kapital],
            ['Основные средства', self.OS, 'Добавочный капитал', self.dobavochniy_kapital],
            [None, None, 'Резервный капитал', self.reservniy_kapital],
            [None, None, 'Нераспределенная прибыль (непокрытый убыток)', self.neraspred_pribil],
            ['Итого по разделу 1', self.r1, 'Итого по разделу 3', self.r3],
            [None, None, None, None],

            ['2. Оборотные активы', None, '4. Долгосрочные обязательства', None],
            ['Запасы', self.r_K_ob_zap, '', None],
            [' сырье и материалы', self.K_ob_sr_pr_zap, 'Итого по 4 разделу', self.r4],
            [' затраты в незавершенном производстве', self.K_ob_nez_pr, '', None],
            [' готовая продукция и товары для перепродажи', self.K_ob_got_prod, '5. Краткосрочные обязательства', None],
            [' расходы будущих периодов', self.K_ob_RBP, 'Заемные средства', self.kratkosroch_zaem_sredstva],
            [' прочие запасы и затраты', self.K_ob_extra, 'Прочие обязательства', self.kratkosroch_prochee],
            ['Дебиторская задолженность', self.debitor_dolg, '', None],
            ['Денежные средства', self.K_ob_ds, '', None],
            [None, None, None, None],
            ['Итого по разделу 2', self.r2, 'Итого по разделу 5', self.r5],
            [None, None, None, None],
            ['Баланс', self.active, 'Баланс', self.passive],
        ]


