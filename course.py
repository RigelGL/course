import texttable
from typing import *


class Value:
    """
    Значение, которое имеет переменные и/или постоянные затрты
    Может иметь потомков
    """
    __slots__ = ['name', '_const', '_variable', 'children']

    def __init__(self, name: str, const: float = 0, variable: float = 0):
        self.name = name
        self._const = const
        self._variable = variable
        self.children = None

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

    def __str__(self, deep=0):
        c = self.const
        v = self.variable
        if c > 0 and v > 0:
            s = '  ' * deep + '{}, постоянные: {:,.2f}, переменные: {:,.2f}, все: {:,.2f}'.format(self.name, c, v, c + v)
        elif c == 0:
            s = '  ' * deep + '{}, переменные: {:,.2f}'.format(self.name, v)
        else:
            s = '  ' * deep + '{}, постоянные: {:,.2f}'.format(self.name, c)
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
    __slots__ = ['name', 'percent']

    def __init__(self, name, percent):
        self.name = name
        self.percent = percent


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

    def add_row(self, name: str, percent: float):
        if name in self._rows_map.keys():
            print('error, double value', name)
            return
        percent /= 100.0
        r = _PerPercentTableRow(name, percent)
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

    def __str__(self):
        t = texttable.Texttable(max_width=200)
        t.set_cols_dtype([str] * 3)
        t.header(['name', 'percent', 'value'])
        total = 0
        for i in self.rows:
            np = i.percent / self._total_percent if self._normalize else i.percent
            v = np * self._initial_value
            if type(self._initial_value) == int:
                v = max(1 if self._minimum_is_one else 0, round(v))
            total += v
            t.add_row([i.name, '{:.2f}'.format(np * 100), ('{:,.2f}' if type(self._initial_value) == float else '{:,}').format(v)])
        t.add_row(['total', 100.0 if self._normalize else self._total_percent,
                   ('{:,.2f}' if type(self._initial_value) == float else '{:,}').format(total)])
        return t.draw()
