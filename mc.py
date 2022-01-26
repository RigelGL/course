import latex2mathml.converter
import numpy as np
import math
from lxml import etree
import matplotlib
from matplotlib import pyplot as plt
from io import BytesIO
from docx import Document
from docx.shared import Inches, Mm, Cm, Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_BREAK
from docx.enum.style import WD_STYLE_TYPE
from course import Table, Value, PerPercentTable, CalculateTable, ActivePassive

# formatting
document = Document()
FONT_NAME = 'Times New Roman'
title_size = Pt(16)
text_size = Pt(14)
first_line_indent = Cm(1.25)

title_text = None
subtitle_text = None
main_text = None
table_name_text = None
formula_style = None
formula_style_12 = None
table_style = None
table_style_12 = None
table_style_12_dense = None
table_style_10 = None


def init_styles():
    global main_text, title_text, table_name_text, subtitle_text, \
        formula_style, formula_style_12, table_style, table_style_12, table_style_12_dense, table_style_10

    main_text = document.styles.add_style('Main text', WD_STYLE_TYPE.PARAGRAPH)
    main_text.paragraph_format.first_line_indent = Cm(1.25)
    main_text.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    main_text.font.name = FONT_NAME
    main_text.font.size = Pt(14)
    main_text.font.bold = False
    main_text.next_paragraph_style = main_text

    subtitle_text = document.styles.add_style('Subtitle text', WD_STYLE_TYPE.PARAGRAPH)
    subtitle_text.base_style = document.styles['Heading 2']
    subtitle_text.font.name = FONT_NAME
    subtitle_text.font.size = Pt(14)
    subtitle_text.font.bold = False
    subtitle_text.paragraph_format.space_after = Pt(8)
    subtitle_text.next_paragraph_style = main_text

    title_text = document.styles.add_style('Title text', WD_STYLE_TYPE.PARAGRAPH)
    title_text.base_style = document.styles['Heading 1']
    title_text.font.name = FONT_NAME
    title_text.font.size = Pt(16)
    title_text.font.bold = False
    title_text.paragraph_format.space_after = Pt(12)
    title_text.next_paragraph_style = subtitle_text

    table_name_text = document.styles.add_style('Table name text', WD_STYLE_TYPE.PARAGRAPH)
    table_name_text.base_style = document.styles['Main text']
    table_name_text.paragraph_format.first_line_indent = Cm(0)
    table_name_text.paragraph_format.space_before = Pt(12)
    table_name_text.paragraph_format.space_after = Pt(4)
    table_name_text.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    table_name_text.font.size = Pt(12)

    formula_style = document.styles.add_style('Formula style', WD_STYLE_TYPE.PARAGRAPH)
    formula_style.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    formula_style.font.name = FONT_NAME
    formula_style.font.size = Pt(14)
    formula_style.font.bold = False

    formula_style_12 = document.styles.add_style('Formula style 12', WD_STYLE_TYPE.PARAGRAPH)
    formula_style_12.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    formula_style_12.font.name = FONT_NAME
    formula_style_12.font.size = Pt(12)
    formula_style_12.font.bold = False

    table_style = document.styles.add_style('Table style', WD_STYLE_TYPE.TABLE)
    table_style.base_style = document.styles['Table Grid']
    table_style.paragraph_format.space_before = Mm(1.5)
    table_style.paragraph_format.space_after = Mm(1.5)
    table_style.font.name = FONT_NAME
    table_style.font.size = Pt(14)
    table_style.next_paragraph_style = main_text

    table_style_12 = document.styles.add_style('Table style 12', WD_STYLE_TYPE.TABLE)
    table_style_12.base_style = document.styles['Table Grid']
    table_style_12.paragraph_format.space_before = Mm(1.5)
    table_style_12.paragraph_format.space_after = Mm(1.5)
    table_style_12.font.name = FONT_NAME
    table_style_12.font.size = Pt(12)
    table_style_12.next_paragraph_style = main_text

    table_style_12_dense = document.styles.add_style('Table style 12 dense', WD_STYLE_TYPE.TABLE)
    table_style_12_dense.base_style = document.styles['Table Grid']
    table_style_12_dense.paragraph_format.left_indent = Mm(0.25)
    table_style_12_dense.paragraph_format.right_indent = Mm(0.25)
    table_style_12_dense.paragraph_format.space_before = Mm(0.15)
    table_style_12_dense.paragraph_format.space_after = Mm(0.15)
    table_style_12_dense.font.name = FONT_NAME
    table_style_12_dense.font.size = Pt(12)
    table_style_12_dense.next_paragraph_style = main_text

    table_style_10 = document.styles.add_style('Table style 10', WD_STYLE_TYPE.TABLE)
    table_style_10.base_style = document.styles['Table Grid']
    table_style_10.paragraph_format.space_before = Mm(1.0)
    table_style_10.paragraph_format.space_after = Mm(1.0)
    table_style_10.font.name = FONT_NAME
    table_style_10.font.size = Pt(10)
    table_style_10.next_paragraph_style = main_text


def dp(text='', style=None):
    if style is None:
        return document.add_paragraph(text, main_text)
    return document.add_paragraph(text, style)


def latex_to_word(latex_input):
    mathml = latex2mathml.converter.convert(latex_input)
    tree = etree.fromstring(mathml)
    xslt = etree.parse('MML2OMML.XSL')
    transform = etree.XSLT(xslt)
    new_dom = transform(tree)
    return new_dom.getroot()


def add_formula(latex, p=None, style=None):
    if p is None:
        if style is None:
            p = dp(style=formula_style)
        else:
            p = dp(style=style)
        p.add_run()._element.append(latex_to_word(latex))
    else:
        p.add_run()._element.append(latex_to_word(latex))
    return p


def add_formula_with_description(latex, description, style=None):
    p = add_formula(latex, style=style)
    fs = Pt(14)
    p.add_run('\nГде:\n').font.size = fs
    for i, e in enumerate(description):
        add_formula(e[0], p)
        p.add_run(' - ' + e[1] + ('\n' if i + 1 < len(description) else '')).font.size = fs
    return p


def add_table(data, widths=None, first_bold=False, style=None):
    table = document.add_table(rows=len(data), cols=len(data[0]))
    if style is None:
        table.style = table_style
    else:
        table.style = style
    for i in range(len(table.rows)):
        row = table.rows[i]
        for j in range(len(row.cells)):
            cell = row.cells[j]
            if data[i][j] is not None:
                r = cell.paragraphs[0].add_run(data[i][j])
                if first_bold and i == 0:
                    r.font.bold = True
            if widths and widths[j] is not None:
                cell.width = widths[j]
    return table


def add_active_passive_table(active_passive):
    table = document.add_table(rows=21, cols=4)
    table.style = table_style_12_dense
    table.autofit = False
    table.allow_autofit = True
    for i, e in enumerate(['АКТИВ', 'руб', 'ПАССИВ', 'руб']):
        table.cell(0, i).paragraphs[0].add_run(e)
        table.cell(0, i).width = Cm(5.75 if i % 2 == 0 else 3.25)

    tbl = active_passive.to_table()

    for i in range(len(tbl)):
        for j in range(4):
            table.cell(i + 1, j).width = Cm(5.5 if j % 2 == 0 else 3.0)
            if tbl[i][j] is not None:
                if type(tbl[i][j]) != str:
                    table.cell(i + 1, j).paragraphs[0].add_run(fn(tbl[i][j]))
                else:
                    table.cell(i + 1, j).paragraphs[0].add_run(tbl[i][j].strip())
                    if tbl[i][j].startswith(' '):
                        table.cell(i + 1, j).paragraphs[0].paragraph_format.left_indent = Cm(0.5)

    for e in [
        [0, 0], [0, 2],
        [1, 0], [1, 2],
        [6, 0], [6, 2],
        [8, 0], [8, 2],
        [10, 2],
        [12, 2],
        [18, 0], [18, 2],
        [20, 0], [20, 2]
    ]:
        table.cell(e[0], e[1]).paragraphs[0].runs[0].bold = True
        table.cell(e[0], e[1]).paragraphs[0].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    return table


def fn(num, ln=2):
    return ('{:,.' + str(ln) + 'f}').format(num)


class InitialData:
    __slots__ = ['N_pl', 'materials', 'accessories', 'operations']

    def __init__(self, N_pl=45_000):
        self.N_pl = N_pl
        self.materials = Table('name', 'cost', 'amount', 't_zap')

        self.materials.add_row('а', 60, 1, 30)
        self.materials.add_row('б', 150, 3, 40)
        self.materials.add_row('в', 350, 3, 60)
        self.materials.add_row('г', 60, 2, 50)

        self.accessories = Table('name', 'cost', 'amount', 't_zap')
        self.accessories.add_row('а', 50, 1, 35)
        self.accessories.add_row('б', 60, 3, 70)
        self.accessories.add_row('в', 70, 2, 45)

        self.operations = Table('cost', 'time', 'name')
        self.operations.add_row(0, 0.4, '-ручная операция-')
        self.operations.add_row(500_000, 0.3, 'а')
        self.operations.add_row(600_000, 0.2, 'б')
        self.operations.add_row(700_000, 0.6, 'в')


class Chapter_1:
    __slots__ = [
        'T_pl', 'B', 'C', 'D', 'O', 'H', 'gamma', 'F_ob_ef',
        'N_machine_work_percent', 'n_ob_k_rasch', 'n_ob_k_fact', 'b_fact',
        'TO_perv', 'main_resources', 'S_os', 'S_os_amortisable']

    def __init__(self, initial_data: InitialData):

        def calc_1_1_n_ob_k(N, t, b, F, ceil=True):
            numerator = N * t
            denominator = b * F
            if ceil:
                return math.ceil(numerator / denominator)
            return numerator / denominator

        self.T_pl = 365
        self.B = 128
        self.C = 2
        self.D = 8
        self.O = 20
        self.H = 20
        self.gamma = 0.05
        self.F_ob_ef = (self.T_pl - self.B) * self.C * self.D * (1 - self.gamma)
        self.N_machine_work_percent = 0.75

        self.n_ob_k_rasch = []
        self.n_ob_k_fact = []
        self.b_fact = []

        self.TO_perv = 0

        for operation in initial_data.operations.rows[1:]:
            e = calc_1_1_n_ob_k(initial_data.N_pl, operation['time'], self.N_machine_work_percent, self.F_ob_ef, False)
            self.n_ob_k_rasch.append(e)
            e = math.ceil(e)
            self.n_ob_k_fact.append(e)
            self.b_fact.append(calc_1_1_n_ob_k(initial_data.N_pl, operation[1], e, self.F_ob_ef, False))
            self.TO_perv += operation['cost'] * e

        self.main_resources = Table('n', 'name', '%', 'cost')
        mr = self.main_resources
        self.S_os = self.TO_perv / 0.4

        main_resources = [
            ['1', 'Земля', 0.14],
            ['2', 'Здания', 0.09],
            ['3', 'Сооружения', 0.07],
            ['4', 'Передаточные устройства', 0.06],
            ['5', 'Машины и оборудование, в т.ч.', 0.50],
            ['', '- технологическое оборудование', 0.40],
            ['', '- нетехнологические машины и оборудование', 0.10],
            ['7', 'Транспортные средства', 0.09],
            ['8', 'Инструменты и технологическая оснастка', 0.04],
            ['9', 'Производственный и хозяйственный инвентарь', 0.01],
        ]

        self.S_os_amortisable = round(self.S_os * (1.0 - 0.14), 2)

        for e in main_resources:
            mr.add_row(e[0], e[1], e[2] * 100, self.S_os * e[2])


class Chapter_2:
    __slots__ = [
        'F_rab_ef',
        'C_opr_mean', 'p_mean',
        'R_opr_raw', 'R_opr', 'FOT_opr', 'FOT_opr_extra', 'opr_extra', 'opr_salary',
        'vpr', 'R_vpr', 'FOT_vpr',
        'sl', 'R_sl', 'FOT_sl',
        'R_ppp', 'FOT',
        'insurance_fee', 'FOT_fee', 'FOT_with_fee',
        'stimulating_salary_percent'
    ]

    def __init__(self, initial_data: InitialData, chapter_1: Chapter_1, const=None):
        self.F_rab_ef = (chapter_1.T_pl - chapter_1.B - chapter_1.O - chapter_1.H) * chapter_1.D

        total_time = initial_data.operations.calculate_sum(lambda x: x['time'])

        self.R_opr_raw = initial_data.N_pl * total_time / self.F_rab_ef
        self.R_opr = int(math.ceil(self.R_opr_raw))

        self.vpr = PerPercentTable(self.R_opr, True, False)
        self.vpr.add_row('Настройщик оборудования', 0.05, 60_000)
        self.vpr.add_row('Складовщик', 0.07, 50_000)
        self.vpr.add_row('Уборщик', 0.05, 30_000)
        self.vpr.add_row('Контролёр ОТК', 0.07, 80_000)

        self.sl = PerPercentTable(self.R_opr, True, False)
        self.sl.add_row('Генеральный директор', 0, 120_000)
        self.sl.add_row('HR', 0, 70_000)
        self.sl.add_row('Менеджер по закупу', 0, 70_000)
        self.sl.add_row('Менеджер по производству', 0, 85_000)
        self.sl.add_row('Инженер', 2. / self.R_opr, 85_000)
        self.sl.add_row('Бухгалтер', 2. / self.R_opr, 80_000)
        self.sl.add_row('Охраниик', 4. / self.R_opr, 35_000)
        self.sl.add_row('Логист', 3. / self.R_opr, 50_000)
        self.sl.add_row('Курьер', 0.5 - 10. / self.R_opr, 45_000)

        self.R_vpr = self.vpr.total
        self.R_sl = self.sl.total

        self.opr_salary = 50_000
        self.opr_extra = 5_000

        self.stimulating_salary_percent = 1.0

        self.C_opr_mean = round(12 * self.opr_salary / self.F_rab_ef, 2)
        self.p_mean = round(self.C_opr_mean * total_time / len(initial_data.operations.rows), 2)

        self.FOT_opr = self.p_mean * initial_data.N_pl * len(initial_data.operations)
        self.FOT_opr_extra = self.R_opr * (self.opr_extra * 12 + self.stimulating_salary_percent * self.opr_salary)

        self.FOT_vpr = self.vpr.calc_sum(lambda amount, salary: amount * salary * (12 + self.stimulating_salary_percent))
        self.FOT_sl = self.sl.calc_sum(lambda amount, salary: amount * salary * (12 + self.stimulating_salary_percent))

        self.R_ppp = self.R_opr + self.R_vpr + self.R_sl
        if const is None or const['fot'] is None:
            self.FOT = Value('fot', self.FOT_vpr + self.FOT_sl, self.FOT_opr + self.FOT_opr_extra, 'Затраты на оплату труда')
        else:
            self.FOT = Value('fot', const['fot'].const, self.FOT_opr + self.FOT_opr_extra, 'Затраты на оплату труда')

        self.insurance_fee = PerPercentTable(self.FOT.total)
        self.insurance_fee.add_row('ОПФ', 0.22)
        self.insurance_fee.add_row('ФОМС', 0.051)
        self.insurance_fee.add_row('ФСС', 0.029)
        self.insurance_fee.add_row('Страхование от несчастных случаев на производстве и профессиональных заболеваний', 0.04)

        self.FOT_fee = Value('fot fee', round(self.FOT.const * 0.34, 2), round(self.FOT.variable * 0.34, 2), 'Страховые взносы')
        self.FOT_with_fee = Value('fot total', display_name='ФОТ')
        self.FOT_with_fee.add_child(self.FOT)
        self.FOT_with_fee.add_child(self.FOT_fee)


class Chapter_3:
    __slots__ = [
        'S_mat_i_comp',
        'main_materials',
        'help_materials_percent',
        'moving_save_percent',
        'move_save_const_percent',
        'inventory_percent',
        'fuel_percent',
        'fuel_tech_percent',
        'fuel_non_tech_percent',

        'OS_amortisation_percent',
        'OS_amortisation',

        'NMA',
        'NMA_amortisation_percent',
        'NMA_amortisation',

        'OS_fix_percent',
        'OS_fix',

        'extra_percent',

        'costs',
        'S_pr_tek_pl'
    ]

    def calc_mat_costs(self, n, fot, fot_safe, const=None):
        pass

    def __init__(self, initial_data: InitialData, chapter_1: Chapter_1, chapter_2: Chapter_2, const=None):
        self.S_mat_i_comp = initial_data.materials.calculate_sum(lambda x: x['cost'] * x['amount']) + initial_data.accessories.calculate_sum(lambda x: x['cost'] * x['amount'])

        self.costs = Value('proizv', display_name='Затраты')
        self.main_materials = Value('material', variable=initial_data.N_pl * self.S_mat_i_comp, display_name='Материальные затраты')
        m_base = self.main_materials.total

        self.main_materials.add_child(Value('material_main', 0, m_base, display_name='Основные материалы'))

        self.help_materials_percent = 0.05
        self.main_materials.add_child(Value('helper', 0, round(m_base * self.help_materials_percent, 2), display_name='Вспомогательные материалы'))

        self.moving_save_percent = 0.12
        self.move_save_const_percent = 0.3
        ms = round(m_base * self.moving_save_percent, 2)
        msc = round(ms * self.move_save_const_percent, 2) if const is None or const['move save'] is None else const['move save'].const
        msv = round(ms - msc, 2) if const is None or const['move save'] is None else round((1 - self.move_save_const_percent) * ms, 2)
        self.main_materials.add_child(Value('move save', msc, msv, display_name='Транспортно-заготовительные расходы'))

        self.inventory_percent = 0.03
        self.main_materials.add_child(
            Value('inventory', const=round(m_base * self.inventory_percent, 2), display_name='Инструменты, инвентарь')
            if const is None or const['inventory'] is None else const['inventory'])

        self.fuel_percent = 0.55
        fuelt = round(m_base * self.fuel_percent, 2)

        self.fuel_tech_percent = 0.7
        self.fuel_non_tech_percent = 1.0 - self.fuel_tech_percent

        fuel_energy_costs = self.main_materials.add_child(Value('fuel total', display_name='Топливо и энергия'))
        ft = fuel_energy_costs.add_child(Value('fuel tech', variable=round(fuelt * self.fuel_tech_percent, 2), display_name='Технологическое топливо и энергия'))
        fuel_energy_costs.add_child(
            Value('fuel non tech', const=fuelt - ft.total, display_name='Нетехнологическое топливо и энергия')
            if const is None or const['fuel non tech'] is None else const['fuel non tech'])
        self.costs.add_child(self.main_materials)

        self.costs.add_child(chapter_2.FOT)
        self.costs.add_child(chapter_2.FOT_fee)

        self.OS_amortisation_percent = 0.1
        self.OS_amortisation = round(self.OS_amortisation_percent * chapter_1.S_os_amortisable, 2)

        self.NMA = 3_000_000
        self.NMA_amortisation_percent = 0.1
        self.NMA_amortisation = round(self.NMA_amortisation_percent * self.NMA, 2)

        aos = Value('amortisation OS', const=self.OS_amortisation, display_name='Амортизация ОС')
        anma = Value('amortisation NMA', const=self.NMA_amortisation, display_name='Амортизация НМА')
        a = Value('amortisation', display_name='Амортизация ОС и НМА')
        a.add_child(aos)
        a.add_child(anma)
        self.costs.add_child(a)

        self.OS_fix_percent = 0.06
        self.OS_fix = round(self.OS_fix_percent * chapter_1.S_os_amortisable, 2)

        self.extra_percent = 0.05
        self.costs.add_child(Value('extra', const=self.costs.const * self.extra_percent + self.OS_fix, variable=self.costs.variable * self.extra_percent, display_name='Прочие затраты'))

        self.S_pr_tek_pl = self.costs.total


class Chapter_4:
    __slots__ = [
        'N_pl_values',
        'S_b_proizv',
        'S_b_poln',
        'S_kom_percent',
        'S_kom_const_percent',
        'S_kom',
        'S_sum',
        'ct1'
    ]

    def calc_n(self, n):
        fake_initial = InitialData(n)
        fake_chapter_2 = Chapter_2(fake_initial, chapter_1, chapter_2.FOT)
        fake_chapter_3 = Chapter_3(fake_initial, chapter_1, fake_chapter_2, const=chapter_3.costs)
        fake_chapter_4 = Chapter_4(fake_initial, fake_chapter_3, self.S_kom)
        return fake_chapter_4.S_sum.head()

    def __init__(self, initial_data: InitialData, chapter_3: Chapter_3, const=None):
        self.N_pl_values = [450, 2700, 7200, 18900, 33750, 45000]
        self.S_b_proizv = chapter_3.S_pr_tek_pl / initial_data.N_pl
        self.S_kom_percent = 0.04
        self.S_kom_const_percent = 0.6
        s = round(self.S_b_proizv * self.S_kom_percent, 0) * initial_data.N_pl
        sc = round(s * self.S_kom_const_percent, 2) if const is None or const['S_kom'] is None else const['S_kom'].const
        sv = round(s - sc, 2) if const is None or const['S_kom'] is None else s * (1 - self.S_kom_const_percent)

        self.S_sum = Value('S_sum', display_name='Суммарные затраты')
        self.S_sum.add_child(chapter_3.costs)
        self.S_kom = Value('S_kom', sc, sv, 'Коммерческие затраты')
        self.S_sum.add_child(self.S_kom)
        self.S_b_poln = Value('S_b_poln', self.S_sum.const / initial_data.N_pl, self.S_sum.variable / initial_data.N_pl)

        if const is None:
            self.ct1 = CalculateTable(self.N_pl_values, Chapter_4.calc_n, self)
        else:
            self.ct1 = None


class Chapter_5:
    __slots__ = [
        'K_ob_sr_mk',
        'k_ob_sr_percent',
        'K_ob_sr_pr_zap',

        'k_nz',
        'gamma_cycle',
        'T_cycle',
        'K_ob_nez_pr',

        't_real',
        'K_ob_got_prod',

        'gamma_ob',
        'K_ob_extra',
        'K_ob_sum',
    ]

    def __init__(self, initial_data: InitialData, chapter_1: Chapter_1, chapter_3: Chapter_3, chapter_4: Chapter_4):
        mz = round(initial_data.materials.calculate_sum(lambda x: x['amount'] * x['cost'] * initial_data.N_pl / chapter_1.T_pl * x['t_zap']), 2)
        cz = round(initial_data.accessories.calculate_sum(lambda x: x['amount'] * x['cost'] * initial_data.N_pl / chapter_1.T_pl * x['t_zap']), 2)
        self.K_ob_sr_mk = round(mz + cz, 2)
        self.k_ob_sr_percent = 0.4
        self.K_ob_sr_pr_zap = round((1 + self.k_ob_sr_percent) * self.K_ob_sr_mk, 2)
        self.k_nz = (chapter_3.S_mat_i_comp + chapter_4.S_b_proizv) / (chapter_4.S_b_proizv * 2)
        self.gamma_cycle = 50

        self.T_cycle = round(initial_data.operations.calculate_sum(lambda x: x['time']) * self.gamma_cycle /
                             (chapter_1.C * chapter_1.D) * chapter_1.T_pl / (chapter_1.T_pl - chapter_1.B), 3)

        self.K_ob_nez_pr = chapter_4.S_b_proizv * initial_data.N_pl / chapter_1.T_pl * self.k_nz * self.T_cycle

        self.t_real = 10
        self.K_ob_got_prod = round(chapter_4.S_b_proizv * initial_data.N_pl / chapter_1.T_pl * self.t_real, 2)

        self.gamma_ob = 0.6
        a = round(self.K_ob_sr_pr_zap + self.K_ob_nez_pr + self.K_ob_got_prod, 2)
        self.K_ob_sum = round(a / self.gamma_ob, 2)
        self.K_ob_extra = round(self.K_ob_sum - a, 2)


class Chapter_6:
    __slots__ = [
        'k_ob_RPB_percent',
        'ustavnoy_capital_percent',
        'doldosroch_zaemn_sredstva_percent',
        'kratkosroch_zaemn_sredstva_percent',
        'active_passive'
    ]

    def __init__(self, chapter_1: Chapter_1, chapter_5: Chapter_5):
        self.active_passive = ActivePassive()

        self.active_passive.NMA = chapter_3.NMA
        self.active_passive.OS = chapter_1.S_os

        self.k_ob_RPB_percent = 0.5
        self.active_passive.K_ob_RBP = round(chapter_5.K_ob_extra * self.k_ob_RPB_percent, 2)
        self.active_passive.K_ob_sr_pr_zap = chapter_5.K_ob_sr_pr_zap
        self.active_passive.K_ob_ds = chapter_5.K_ob_sum - (chapter_5.K_ob_sr_pr_zap + self.active_passive.K_ob_RBP)

        self.ustavnoy_capital_percent = 0.8
        self.active_passive.ustavnoy_kapital = round(self.active_passive.active * self.ustavnoy_capital_percent, 2)

        S_summ_passiv_left = self.active_passive.active - self.active_passive.ustavnoy_kapital

        self.doldosroch_zaemn_sredstva_percent = 0.6
        self.active_passive.doldosroch_zaemn_sredstva = round(S_summ_passiv_left * self.doldosroch_zaemn_sredstva_percent, 2)

        self.kratkosroch_zaemn_sredstva_percent = 0.25
        self.active_passive.kratkosroch_zaem_sredstva = round(S_summ_passiv_left * self.kratkosroch_zaemn_sredstva_percent, 2)
        self.active_passive.kratkosroch_prochee = round(S_summ_passiv_left - self.active_passive.doldosroch_zaemn_sredstva - self.active_passive.kratkosroch_zaem_sredstva, 2)


class Chapter_7:
    __slots__ = [
        'tax',
        'net_profit_percent',
        'net_profit',
        'profit_before_tax',
        'k_nats',
        'P_b_poln',
        'P_b_perem',
        'P_proizv_plan',
        'price_fact_percent',
        'P_fact'
    ]

    def __init__(self, chapter_4: Chapter_4, chapter_6: Chapter_6):
        self.tax = 0.2
        self.net_profit_percent = 0.6
        self.net_profit = round(chapter_6.active_passive.ustavnoy_kapital * self.net_profit_percent, 2)
        self.profit_before_tax = round(self.net_profit / (1 - self.tax), 2)
        self.k_nats = self.profit_before_tax / chapter_4.S_sum.total
        self.P_b_poln = round(chapter_4.S_b_poln.total * (1 + self.k_nats), 2)
        self.P_b_perem = round(chapter_4.S_b_poln.variable * (1 + (self.profit_before_tax + chapter_4.S_sum.const) / chapter_4.S_sum.variable), 2)
        self.P_proizv_plan = max(self.P_b_poln, self.P_b_perem)
        self.price_fact_percent = 0.94
        self.P_fact = round(self.P_proizv_plan * self.price_fact_percent, 2)


class Chapter_8:
    __slots__ = [
        'N_fact_percent',
        'N_fact',
        'N_ost',
        'Q_plan', 'Q_fact',
        'K_ob_got_prod_plan', 'K_ob_got_prod_fact',
        'S_pr_got_pr_plan', 'S_pr_got_pr_fact',
        'S_valovaya_plan', 'S_valovaya_fact',
        'kom_percent',
        'S_kom_plan', 'S_kom_fact',
        'P_pr_plan', 'P_pr_fact',
        'P_pr_do_nalogov_plan', 'P_pr_do_nalogov_fact',
        'pr_dir_fact_percent',
        'S_prochie_dohidy_i_rashody_plan', 'S_prochie_dohidy_i_rashody_fact',
        'nalog_na_pribil_plan', 'nalog_na_pribil_fact',
        'P_chistaya_plan', 'P_chistaya_fact'
    ]

    def __init__(self, initial_data: InitialData, chapter_3: Chapter_3, chapter_4: Chapter_4, chapter_5: Chapter_5, chapter_7: Chapter_7):
        self.N_fact_percent = 0.95
        self.N_fact = int(self.N_fact_percent * initial_data.N_pl)
        self.N_ost = initial_data.N_pl - self.N_fact

        self.Q_plan = chapter_7.P_proizv_plan * initial_data.N_pl
        self.Q_fact = chapter_7.P_fact * self.N_fact

        self.K_ob_got_prod_plan = chapter_5.K_ob_got_prod
        self.K_ob_got_prod_fact = chapter_5.K_ob_got_prod + round(chapter_4.S_b_proizv * self.N_ost, 2)
        self.S_pr_got_pr_plan = chapter_3.S_pr_tek_pl - chapter_5.K_ob_nez_pr - self.K_ob_got_prod_plan
        self.S_pr_got_pr_fact = chapter_3.S_pr_tek_pl - chapter_5.K_ob_nez_pr - self.K_ob_got_prod_fact

        self.S_valovaya_plan = self.Q_plan - self.S_pr_got_pr_plan
        self.S_valovaya_fact = self.Q_fact - self.S_pr_got_pr_fact

        self.kom_percent = 0.94
        self.S_kom_plan = chapter_4.S_kom.total
        self.S_kom_fact = self.kom_percent * chapter_4.S_kom.total

        self.P_pr_plan = self.S_valovaya_plan - self.S_kom_plan
        self.P_pr_fact = self.S_valovaya_fact - self.S_kom_fact

        self.P_pr_do_nalogov_plan = chapter_7.profit_before_tax
        self.pr_dir_fact_percent = 0.93
        self.S_prochie_dohidy_i_rashody_plan = self.P_pr_plan - self.P_pr_do_nalogov_plan
        self.S_prochie_dohidy_i_rashody_fact = round(self.S_prochie_dohidy_i_rashody_plan * self.pr_dir_fact_percent, 2)
        self.P_pr_do_nalogov_fact = self.P_pr_fact - self.S_prochie_dohidy_i_rashody_fact
        self.nalog_na_pribil_plan = self.P_pr_do_nalogov_plan - chapter_7.net_profit
        self.nalog_na_pribil_fact = round(self.P_pr_do_nalogov_fact * chapter_7.tax, 2)

        self.P_chistaya_plan = chapter_7.net_profit
        self.P_chistaya_fact = self.P_pr_do_nalogov_fact - self.nalog_na_pribil_fact


class Chapter_9:
    __slots__ = [
        'K_den_sr_plan', 'K_den_sr_fact',
        'K_den_sr_konez_plan', 'K_den_sr_konez_fact',
        'S_kratkosroch_zaem_sredstva_konez_plan',
        'S_kratkosroch_zaem_sredstva_konez_fact',
        'active_passive_plan', 'active_passive_fact',
        'valid_to_cope_kz_plan', 'valid_to_cope_kz_fact',
    ]

    def __init__(self):
        amortisation = chapter_3.costs['amortisation'].total

        self.K_den_sr_plan = chapter_6.active_passive.K_ob_ds + amortisation + chapter_8.P_chistaya_plan - (chapter_5.K_ob_nez_pr + chapter_8.K_ob_got_prod_plan)
        if self.K_den_sr_plan - 500_000 > chapter_6.active_passive.kratkosroch_zaem_sredstva:
            self.K_den_sr_konez_plan = self.K_den_sr_plan - chapter_6.active_passive.kratkosroch_zaem_sredstva
            self.S_kratkosroch_zaem_sredstva_konez_plan = 0
            self.valid_to_cope_kz_plan = 'full'
        elif self.K_den_sr_plan > 500_000:
            self.K_den_sr_konez_plan = 500_000
            self.S_kratkosroch_zaem_sredstva_konez_plan = chapter_6.active_passive.kratkosroch_zaem_sredstva - (self.K_den_sr_plan - 500_000)
            self.valid_to_cope_kz_plan = 'part'
        else:
            self.S_kratkosroch_zaem_sredstva_konez_plan = chapter_6.active_passive.kratkosroch_zaem_sredstva
            self.K_den_sr_konez_plan = self.K_den_sr_plan
            self.valid_to_cope_kz_plan = 'none'

        self.K_den_sr_fact = chapter_6.active_passive.K_ob_ds + amortisation + chapter_8.P_chistaya_fact - (chapter_5.K_ob_nez_pr + chapter_8.K_ob_got_prod_fact)
        if self.K_den_sr_fact - 500_000 > chapter_6.active_passive.kratkosroch_zaem_sredstva:
            self.K_den_sr_konez_fact = self.K_den_sr_fact - chapter_6.active_passive.kratkosroch_zaem_sredstva
            self.S_kratkosroch_zaem_sredstva_konez_fact = 0
            self.valid_to_cope_kz_fact = 'full'
        elif self.K_den_sr_fact > 500_000:
            self.K_den_sr_konez_fact = 500_000
            self.S_kratkosroch_zaem_sredstva_konez_fact = chapter_6.active_passive.kratkosroch_zaem_sredstva - (self.K_den_sr_fact - 500_000)
            self.valid_to_cope_kz_fact = 'part'
        else:
            self.S_kratkosroch_zaem_sredstva_konez_fact = chapter_6.active_passive.kratkosroch_zaem_sredstva
            self.K_den_sr_konez_fact = self.K_den_sr_fact
            self.valid_to_cope_kz_fact = 'none'

        p = self.active_passive_plan = ActivePassive()
        f = self.active_passive_fact = ActivePassive()

        f.NMA = p.NMA = chapter_3.NMA - chapter_3.costs['amortisation NMA'].total
        f.OS = p.OS = chapter_1.S_os - chapter_3.costs['amortisation OS'].total

        f.K_ob_sr_pr_zap = p.K_ob_sr_pr_zap = chapter_5.K_ob_sr_pr_zap
        f.K_ob_nez_pr = p.K_ob_nez_pr = chapter_5.K_ob_nez_pr
        p.K_ob_got_prod = chapter_5.K_ob_got_prod
        p.K_ob_RBP = chapter_6.active_passive.K_ob_RBP
        p.K_ob_ds = self.K_den_sr_konez_plan

        p.ustavnoy_kapital = chapter_6.active_passive.ustavnoy_kapital
        p.neraspred_pribil = chapter_8.P_chistaya_plan

        p.doldosroch_zaemn_sredstva = chapter_6.active_passive.doldosroch_zaemn_sredstva

        p.kratkosroch_zaem_sredstva = 0
        p.kratkosroch_prochee = chapter_6.active_passive.kratkosroch_prochee

        f.K_ob_got_prod = chapter_8.K_ob_got_prod_fact
        f.K_ob_RBP = chapter_6.active_passive.K_ob_RBP
        f.K_ob_ds = self.K_den_sr_konez_fact

        f.ustavnoy_kapital = chapter_6.active_passive.ustavnoy_kapital
        f.neraspred_pribil = chapter_8.P_chistaya_fact

        f.doldosroch_zaemn_sredstva = chapter_6.active_passive.doldosroch_zaemn_sredstva

        f.kratkosroch_zaem_sredstva = self.S_kratkosroch_zaem_sredstva_konez_fact
        f.kratkosroch_prochee = chapter_6.active_passive.kratkosroch_prochee


class Chapter_10:
    __slots__ = [
        'k_sob_ob_sr_plan', 'k_sob_ob_sr_fact',
        'k_obespech_sob_sr_plan', 'k_obespech_sob_sr_fact',
        'k_abs_likvid_plan', 'k_abs_likvid_fact',
        'k_tek_likvid_plan', 'k_tek_likvid_fact',
        'V',
        'OS_year_mean',
        'k_FO_plan', 'k_FO_fact',
        'k_FE_plan', 'k_FE_fact',
        'Z_ob_sr_year_mean_plan', 'Z_ob_sr_year_mean_fact',
        'Z_ob_plan', 'Z_ob_fact',
        'S_sobstv_cap_year_mean_plan', 'S_sobstv_cap_year_mean_fact',
        'k_oborach_sobstv_capital_plan', 'k_oborach_sobstv_capital_fact',
        'R_production_plan', 'R_production_fact',
        'R_sell_plan', 'R_sell_fact',
        'R_active_plan', 'R_active_fact',
        'R_sobstv_capital_plan', 'R_sobstv_capital_fact',

        'N_kr',
        'Q_kr',
        'k_pokr',
        'Q_fin_pr_plan', 'Q_fin_pr_fact',
        'proizv_richag_plan', 'proizv_richag_fact'
    ]

    @staticmethod
    def calc_n(n):
        fake_initial = InitialData(n)
        fake_chapter_2 = Chapter_2(fake_initial, chapter_1, chapter_2.FOT)
        fake_chapter_3 = Chapter_3(fake_initial, chapter_1, fake_chapter_2, const=chapter_3.costs)
        fake_chapter_4 = Chapter_4(fake_initial, fake_chapter_3, chapter_4.S_kom)
        return fake_chapter_4.S_sum

    @staticmethod
    def bin_search():
        N_kr_left = 0
        N_kr_right = initial_data.N_pl

        while N_kr_right - N_kr_left > 1:
            N_kr_mean = round((N_kr_right + N_kr_left) / 2)

            s = Chapter_10.calc_n(N_kr_mean)

            if N_kr_mean * chapter_7.P_proizv_plan > s['S_sum'].total:
                N_kr_right = N_kr_mean
            else:
                N_kr_left = N_kr_mean

        return N_kr_left

    @staticmethod
    def calc_k_pokr(n):
        fake_initial = InitialData(n)
        fake_chapter_2 = Chapter_2(fake_initial, chapter_1, const=chapter_2.FOT)
        fake_chapter_3 = Chapter_3(fake_initial, chapter_1, fake_chapter_2, const=chapter_3.costs)
        fake_chapter_4 = Chapter_4(fake_initial, fake_chapter_3, const=chapter_4.S_kom)

        k_pokr = (chapter_7.P_proizv_plan - fake_chapter_4.S_b_poln.variable) / chapter_7.P_proizv_plan

        return k_pokr

    def __init__(self, initial_data: InitialData,
                 chapter_1: Chapter_1, chapter_2: Chapter_2, chapter_3: Chapter_3, chapter_4: Chapter_4,
                 chapter_6: Chapter_6, chapter_7: Chapter_7, chapter_8: Chapter_8, chapter_9: Chapter_9):
        self.k_sob_ob_sr_plan = chapter_9.active_passive_plan.r2 - chapter_9.active_passive_plan.r5
        self.k_sob_ob_sr_fact = chapter_9.active_passive_fact.r2 - chapter_9.active_passive_fact.r5

        self.k_obespech_sob_sr_plan = self.k_sob_ob_sr_plan / chapter_9.active_passive_plan.r2
        self.k_obespech_sob_sr_fact = self.k_sob_ob_sr_fact / chapter_9.active_passive_fact.r2

        self.k_abs_likvid_plan = chapter_9.active_passive_plan.K_ob_ds / chapter_9.active_passive_plan.r5
        self.k_abs_likvid_fact = chapter_9.active_passive_fact.K_ob_ds / chapter_9.active_passive_fact.r5

        self.k_tek_likvid_plan = chapter_9.active_passive_plan.r2 / chapter_9.active_passive_plan.r5
        self.k_tek_likvid_fact = chapter_9.active_passive_fact.r2 / chapter_9.active_passive_fact.r5

        self.V = initial_data.N_pl / chapter_2.R_ppp

        self.OS_year_mean = round(chapter_1.S_os_amortisable - chapter_3.costs['amortisation OS'].total * 0.5, 2)
        self.k_FO_plan = chapter_8.Q_plan / self.OS_year_mean
        self.k_FO_fact = chapter_8.Q_fact / self.OS_year_mean

        self.k_FE_plan = 1 / self.k_FO_plan
        self.k_FE_fact = 1 / self.k_FO_fact

        self.Z_ob_sr_year_mean_plan = round((chapter_6.active_passive.r2 + chapter_9.active_passive_plan.r2) * 0.5, 2)
        self.Z_ob_sr_year_mean_fact = round((chapter_6.active_passive.r2 + chapter_9.active_passive_fact.r2) * 0.5, 2)
        self.Z_ob_plan = chapter_8.Q_plan / self.Z_ob_sr_year_mean_plan
        self.Z_ob_fact = chapter_8.Q_fact / self.Z_ob_sr_year_mean_fact

        self.S_sobstv_cap_year_mean_plan = round((chapter_6.active_passive.r3 + chapter_9.active_passive_plan.r3) * 0.5, 2)
        self.S_sobstv_cap_year_mean_fact = round((chapter_6.active_passive.r3 + chapter_9.active_passive_fact.r3) * 0.5, 2)
        self.k_oborach_sobstv_capital_plan = chapter_8.Q_plan / self.S_sobstv_cap_year_mean_plan
        self.k_oborach_sobstv_capital_fact = chapter_8.Q_fact / self.S_sobstv_cap_year_mean_fact

        self.R_production_plan = chapter_8.P_pr_plan / chapter_4.S_sum.total
        self.R_production_fact = chapter_8.P_pr_fact / chapter_4.S_sum.total

        self.R_sell_plan = chapter_8.P_chistaya_plan / chapter_8.Q_plan
        self.R_sell_fact = chapter_8.P_chistaya_fact / chapter_8.Q_fact

        self.R_active_plan = chapter_8.P_chistaya_plan / chapter_9.active_passive_plan.active
        self.R_active_fact = chapter_8.P_chistaya_fact / chapter_9.active_passive_fact.active

        self.R_sobstv_capital_plan = chapter_8.P_chistaya_plan / self.S_sobstv_cap_year_mean_plan
        self.R_sobstv_capital_fact = chapter_8.P_chistaya_fact / self.S_sobstv_cap_year_mean_fact

        self.N_kr = Chapter_10.bin_search()
        self.Q_kr = self.N_kr * chapter_7.P_proizv_plan
        self.k_pokr = CalculateTable(chapter_4.N_pl_values, Chapter_10.calc_k_pokr)

        self.Q_fin_pr_plan = (chapter_8.Q_plan - self.Q_kr) / chapter_8.Q_plan
        self.Q_fin_pr_fact = (chapter_8.Q_fact - self.Q_kr) / chapter_8.Q_fact

        self.proizv_richag_plan = (chapter_8.Q_plan - chapter_4.S_sum.variable) / chapter_8.P_pr_plan
        self.proizv_richag_fact = (chapter_8.Q_fact - chapter_4.S_b_poln.variable * chapter_8.N_fact) / chapter_8.P_pr_fact


initial_data = InitialData()
chapter_1 = Chapter_1(initial_data)
chapter_2 = Chapter_2(initial_data, chapter_1)
chapter_3 = Chapter_3(initial_data, chapter_1, chapter_2)
chapter_4 = Chapter_4(initial_data, chapter_3)
chapter_5 = Chapter_5(initial_data, chapter_1, chapter_3, chapter_4)
chapter_6 = Chapter_6(chapter_1, chapter_5)
chapter_7 = Chapter_7(chapter_4, chapter_6)
chapter_8 = Chapter_8(initial_data, chapter_3, chapter_4, chapter_5, chapter_7)
chapter_9 = Chapter_9()
chapter_10 = Chapter_10(initial_data, chapter_1, chapter_2, chapter_3, chapter_4, chapter_6, chapter_7, chapter_8, chapter_9)


def gen_introduction():
    document.add_paragraph('Введение', style=title_text)

    paragraphs = [
        'Рассматривается деятельность условного предприятия на протяжение двух периодов хозяйственной деятельности.',

        'В первом разделе условное предприятие начинает свою деятельность с производства лишь одного изделия Б. '
        'Рыночные условия формируются исполнителем путем задания фактических объемов продаж и цен, отличных от плановых значений.',

        'В первом периоде на базе исходных данных производится расчёт основных и оборотных средств; численности персонала и фонда оплаты труда. '
        'Происходит формирование сметы затрат на производство, после чего рассчитывается полная себестоимость продукции. Потом составляется '
        'баланс хозяйственных средств предприятия на начало периода. Рассчитываются плановая цена и, имеющая место на реальном рынке, фактическая '
        'цена. Составляется отчет о прибылях и убытках по плановым и фактическим данным и плановый и фактический баланса хозяйственных средств на '
        'конец периода. По завершении года рассчитываются показатели хозяйственной деятельности и делаются выводы об эффективности работы предприятия.',

        'Во втором разделе предполагается, что анализ рыночного спроса показал, что часть потребителей не удовлетворена качеством выпускаемого изделия '
        'Б. Причем одна группа потребителей готова приобретать аналог А более высокого качества даже по более высокой цене; другая группа потребителей '
        'готова приобретать изделие-аналог В более низкого качества и за более умеренную цену. При этом основной задачей остаётся определить, '
        'в каком количестве производить различные виды продукции.',

        'Во втором разделе также производятся плановые расчёты, аналогичные первому пункту, производится расчёт показателей хозяйственной деятельности '
        'и итогам анализа делаются выводы об эффективности деятельности предприятия и прогнозируется его дальнейшая деятельность.'
    ]

    lr = None
    for text in paragraphs:
        lr = document.add_paragraph(text, style=main_text).runs[0]

    lr.add_break(WD_BREAK.PAGE)


def gen_initial_data():
    def build_consumable_table(name, data):
        table = add_table(
            [[name, 'Стоимость, руб./ед.изм', 'Норма расхода, шт.']] +
            [[str(i + 1), str(e['cost']), str(e['amount'])] for i, e in enumerate(data.rows)],
            [Cm(5), Cm(6), Cm(5)], True)
        table_last_row = table.add_row()
        table_last_row.cells[0].paragraphs[0].add_run('Итого, руб.').bold = True
        table_last_row.cells[1].merge(table_last_row.cells[2])
        table_last_row.cells[1].paragraphs[0].add_run(str(data.calculate_sum(lambda x: x['cost'] * x['amount']))).bold = True
        table_last_row.cells[1].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    dp('Исходные данные 623', title_text)

    dp('Таблица 1, объём производства изделия Б', table_name_text)
    add_table([
        ['Объём производства издения Б, тыс.шт./год', str(initial_data.N_pl // 1000)]
    ], [Cm(11), Cm(5)])
    dp()

    dp('Таблица 2, потребность изделия Б в материалах', table_name_text)
    build_consumable_table('Вид материала', initial_data.materials)
    dp()

    dp('Таблица 3, потребность изделия Б в комплектующих', table_name_text)
    build_consumable_table('Вид комплектующих изделий', initial_data.accessories)
    dp()

    document.add_page_break()

    dp('Таблица 4, технологическая трудоёмкость изделия Б', table_name_text)
    table_4 = add_table(
        [['Номер технологической операции', 'Используемое оборудование', 'Первоначальная стоимость, тыс.руб./ед', 'Технологическая трудоёмкость, час./шт.']] +
        [[str(i + 1), e['name'], str(e['cost'] // 1000), str(e['time'])] for i, e in enumerate(initial_data.operations.rows)], first_bold=True)
    table_4_lr = table_4.add_row()
    table_4_lr.cells[0].merge(table_4_lr.cells[2])
    table_4_lr.cells[0].paragraphs[0].add_run('Итого, час.').bold = True
    table_4_lr.cells[3].paragraphs[0].add_run(str(initial_data.operations.calculate_sum(lambda x: x['time']))).bold = True
    table_4_lr.cells[3].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    dp()


def gen_1_1():
    chapter = chapter_1
    dp('1. Расчет потребности в основных средствах', title_text)
    dp('1.1. Расчет потребности в технологическом оборудовании', subtitle_text)
    dp('Найдём эффективный фонд времени работы оборудования')
    add_formula_with_description('F_{об.эф.} = (A - B) \\cdot C \\cdot D \\cdot (1-γ)', [
        ['A', f'число календарных дней в рассматриваемом периоде = {chapter.T_pl}'],
        ['B', f'число выходных и праздничных дней (выбрать в соответствии с рассматриваемым периодом) = {chapter.B}'],
        ['C', f'число смен в сутки = {chapter.C}'],
        ['D', f'число часов в одной смене = {chapter.D}'],
        ['γ', f'планируемые простои оборудования в долях единицы = {chapter.gamma}'],
    ])

    add_formula(
        f'F_{{об.эф.}} = ({chapter.T_pl} - {chapter.B}) \\cdot {chapter.C} \\cdot '
        f'{chapter.D} \\cdot (1-{chapter.gamma}) = {fn(chapter.F_ob_ef)} час./период').add_run().add_break(WD_BREAK.PAGE)

    dp('После этого найдём само число оборудования', main_text)
    add_formula_with_description('n_{об.k_{расч.}} = \\frac{N_{пл} t_j}{\\beta_{норм} F_{об.эф.}}', [
        ['t_j', 'время обработки изделия или услуги на i-том оборудовании, час/шт'],
        ['N_{пл}', 'планируемый объем производства в рассматриваемом периоде, шт./период'],
        ['\\beta', 'коэффициент загрузки оборудования, 0.75'],
        ['F_{об.эф.}', 'эффективный фонд времени работы оборудования'],
    ])

    dp('Таблица 1.1, Расчет потребности в технологическом оборудовании', table_name_text)
    add_table([
                  ['Операция', 'Расчётное число оборудования', 'Принятое число оборудования', 'Фактический коэффициент нагрузки оборудования']
              ] + [[e[0]['name'], fn(e[1]), str(e[2]), fn(e[3])] for e in zip(initial_data.operations.rows[1:], chapter.n_ob_k_rasch, chapter.n_ob_k_fact, chapter.b_fact)],
              [Cm(4), Cm(4), Cm(4), Cm(4)], True)
    dp()
    dp(
        'Следует отметить, что можно было взять 3 единицы оборудования для операции «в», однако в таком случае '
        'фактический коэффициент нагрузки возрос бы до 0.83, что негативно бы сказалось на сроке службы оборудования. '
        'Поэтому заложим большую надёжность и возьмём 4 единицы.')

    dp('Найдём суммарную первоначальную стоимость технологического оборудования, [тыс. руб.]')

    sm = []
    for i in zip(initial_data.operations.rows[1:], chapter.n_ob_k_fact):
        sm.append(f"{i[0]['cost'] // 1000} \\cdot {i[1]}")

    add_formula('ТО_{перв} = \\sum^{m}_{i}{ТО_{перв i} n_{об i_{прин}}} = ' + ' + '.join(sm) + f' = {fn(chapter.TO_perv // 1000)} [тыс.руб.]', style=formula_style_12)
    document.add_page_break()

    dp('1.2 Стоимостная структура основных средств', title_text)
    dp('Таблица 1.2, стоимостная структура основных средств', table_name_text)
    table = add_table(
        [['№', 'Название', '%', 'Стоимость, руб.']] + [[str(e[0]), e[1], fn(e[2], 0), fn(e[3], 0)] for e in chapter_1.main_resources.rows],
        [Cm(1), Cm(9.5), Cm(1.5), Cm(5)], True
    )
    tr = table.add_row()
    tr.cells[0].merge(tr.cells[1])
    tr.cells[0].paragraphs[0].add_run('Итого').bold = True
    tr.cells[2].paragraphs[0].add_run('100').bold = True
    tr.cells[3].paragraphs[0].add_run(fn(chapter_1.S_os, 0)).bold = True
    dp()
    document.add_page_break()


def gen_1_2():
    dp('2. Расчет численности и фонда з/п персонала предприятия', title_text)
    dp('2.1 Расчет численности ППП', subtitle_text)
    dp('Условно примем что непромышленный персонал отсутствует. Тогда численность ППП складывается из '
       'двух категорий – рабочих и служащих. Рабочие подразделяются на ОПР и ВПР.')
    dp('Чтобы найти численность ОПР найдём в начале эффективный фонд времени одного работающего:')
    add_formula_with_description(
        'F_{раб_{эф}}=(T_{пл}-B-O-H) \\cdot D = ' +
        f"({chapter_1.T_pl}-{chapter_1.B}-{chapter_1.O}-{chapter_1.H}) \\cdot {chapter_1.D} = {chapter_2.F_rab_ef}\\ [час/год\\ чел.]",
        [
            ['O', f'продолжительность отпуска = {chapter_1.H} раб. дн.'],
            ['H', f'число планируемых невыходов = {chapter_1.O} раб. дн.'],
        ],
        style=formula_style_12)
    dp('Примем величину планируемых невыходов за 20 (производство изделий на токарных станках негативно сказывается на здоровье за счёт мелкодисперсной '
       'металлической стружки, переменных магнитных полей электродвигателей и так далее), отпуск 20 рабочих дней (28 – 8 нерабочих дней за 4 недели).')

    dp('Подставим полученное значение в формулу численности ОПР:')

    add_formula_with_description(
        'R_{ОПР}=\\frac{N_{пл} \\sum_{i}^{m}{t_{техн i}}}{F_{раб_{эф}} k_{вн}} = \\frac{' +
        f'{initial_data.N_pl} \\cdot ({"+".join([str(e["time"]) for e in initial_data.operations.rows])})}}{{ {chapter_2.F_rab_ef} \\cdot 1 }} = {chapter_2.R_opr} [чел.]',
        [['t_{техн i}', 'трудоёмкость i-той операции']], style=formula_style_12
    )

    dp()
    p = dp(f'Численность ВРП примем за {fn(chapter_2.R_vpr / chapter_2.R_opr)} R')
    p.add_run('ОПР').font.subscript = True
    p.add_run(', а численность служащих за 0.6R')
    p.add_run('ОПР').font.subscript = True
    p.add_run(':')

    add_formula('R_{ВПР} = ' + fn(chapter_2.R_vpr / chapter_2.R_opr) + 'R_{ОПР} = ' + f'{chapter_2.R_vpr}')
    add_formula('R_{СЛ} = ' + fn(chapter_2.R_sl / chapter_2.R_opr) + 'R_{ОПР} = ' + f'{chapter_2.R_sl}')

    dp('Численность ППП:')
    p = add_formula('R_{ППП} = R_{ОПР} + R_{ВПР} + R_{СЛ} = ' + f'{chapter_2.R_opr} + {chapter_2.R_vpr} + {chapter_2.R_sl} = {chapter_2.R_ppp}')
    p.runs[0].add_break(WD_BREAK.PAGE)

    dp('2.2 Формирование штатного расписания', subtitle_text)
    dp(f'Рабочие ручной операции и весь обслуживающий персонал технологических операций «б», «в», «г» работают по сдельно-премиальной системе. '
       f'При выполнении ими плана, этим работникам положена надбавка в размере {fn(chapter_2.opr_extra, 0)} рублей в месяц.')
    dp('Остальной персонал работает по повременной оплате труда.')
    dp(f'Все работники получают стимулирующие выплаты раз в год в размере {fn(chapter_2.stimulating_salary_percent * 100)}% заработной платы.')

    dp(f'Средняя тарифная ставка ОПР: {fn(chapter_2.opr_salary)} [руб./чел.мес.]')
    add_formula('C_{тар.ст.} =\\frac{12 L_{ОПР ст}}{F_{раб_{эф}}} = \\frac{12 * ' + f'{fn(chapter_2.opr_salary, 0)} }}{{ {chapter_2.F_rab_ef} }} = {fn(chapter_2.C_opr_mean)}')
    add_formula('P_{ср} = C_{тар.ст.} \\cdot t_{ср} = ' + f'{fn(chapter_2.C_opr_mean)}' +
                '\\cdot \\frac{' + '+'.join([str(e['time']) for e in initial_data.operations.rows]) + '}{' + str(len(initial_data.operations)) + f'}} = {fn(chapter_2.p_mean)}')

    document.add_page_break()

    dp('Таблица 2.2.1, состав, структура и заработная плата персонала', table_name_text)
    table = add_table(
        [
            [
                'Должность',
                'Число сотрудников',
                'Тарифная ставка, руб./мес.',
                'Надбавки, руб./мес.',
                'Стимулирующие выплаты, руб./мес.',
                'Категория исполнителей',
                'Система оплаты',
            ]
        ] +
        [[e.name, str(e.amount), fn(e.data, 0), 0, fn(round(e.data / 12), 0), '', ''] for e in chapter_2.sl.rows] +
        [[e.name, str(e.amount), fn(e.data, 0), 0, fn(round(e.data / 12), 0), '', ''] for e in chapter_2.vpr.rows] +
        [['Рабочий', str(chapter_2.R_opr), fn(chapter_2.opr_salary), fn(chapter_2.opr_extra, 0),
          fn(round(chapter_2.opr_salary / 12), 0), 'ОПР', 'Сдельная']],
        [Cm(3.7), Cm(1.75), Cm(2.5), Cm(2.7), Cm(3), Cm(2.0), Cm(1.9)],
        style=table_style_12
    )

    table.cell(1, 5).merge(table.cell(len(chapter_2.sl), 5))
    table.cell(1, 5).paragraphs[0].add_run('Служащие')

    table.cell(1 + len(chapter_2.sl), 5).merge(table.cell(1 + len(chapter_2.sl) + len(chapter_2.vpr) - 1, 5))
    table.cell(1 + len(chapter_2.sl), 5).paragraphs[0].add_run('ВПР')

    table.cell(1, 6).merge(table.cell(1 + len(chapter_2.sl) + len(chapter_2.vpr) - 1, 6))
    table.cell(1, 6).paragraphs[0].add_run('Повременная')
    document.add_page_break()

    dp('Таблица 2.2.2, суммарные заработные платы персонала за год', table_name_text)
    table = add_table(
        [
            [
                'Должность',
                'Число сотрудников',
                'Тарифная ставка, руб./год',
                'Надбавки, руб./год',
                'Стимулирующие выплаты, руб./год',
                'Итого, руб./год'
            ]
        ] +
        [[e.name, str(e.amount), fn(e.data * 12, 0), 0, fn(e.data, 0), fn(e.amount * e.data * 13)] for e in chapter_2.sl.rows] +
        [[e.name, str(e.amount), fn(e.data * 12, 0), 0, fn(e.data, 0), fn(e.amount * e.data * 13)] for e in chapter_2.vpr.rows] +
        [['Рабочий', str(chapter_2.R_opr), fn(chapter_2.opr_salary * 12, 0), fn(chapter_2.opr_extra * 12, 0),
          fn(chapter_2.opr_salary, 0), fn(chapter_2.FOT_opr + chapter_2.FOT_opr_extra)]],
        [Cm(3.7), Cm(2), Cm(2.15), Cm(2.7), Cm(3), Cm(2.25), Cm(1.9)],
        style=table_style_12
    )
    r = table.add_row()
    r.cells[0].paragraphs[0].add_run('Итого').bold = True
    r.cells[1].paragraphs[0].add_run(fn(chapter_2.sl.calc_sum(lambda e, _: e) + chapter_2.vpr.calc_sum(lambda e, _: e) + chapter_2.R_opr, 0)).bold = True
    r.cells[2].paragraphs[0].add_run(fn(chapter_2.sl.calc_sum(lambda e, c: e * c * 12) + chapter_2.vpr.calc_sum(lambda e, c: e * c * 12) + chapter_2.FOT_opr)).bold = True
    r.cells[3].paragraphs[0].add_run(fn(chapter_2.R_opr * chapter_2.opr_extra * 12)).bold = True
    r.cells[4].paragraphs[0].add_run(
        fn(chapter_2.stimulating_salary_percent * (
                chapter_2.sl.calc_sum(lambda e, c: e * c) +
                chapter_2.vpr.calc_sum(lambda e, c: e * c) +
                chapter_2.R_opr * chapter_2.opr_salary
        ))).bold = True
    r.cells[5].paragraphs[0].add_run(fn(chapter_2.FOT.total)).bold = True

    dp()
    dp(f'Отметим, что ОПР получают сдельную зарплату за {fn(chapter_2.R_opr_raw)} чел., но на предприятии их трудится {chapter_2.R_opr}, поэтому каждый из них получит меньше, '
       f'но их суммарная з/п будет равна ФОТ ОПР.').add_run().add_break(WD_BREAK.PAGE)

    dp('2.3 Расчет фонда оплаты труда ППП, величины страховых взносов', subtitle_text)
    dp('Найдём ФОТ ОПР на год:')
    add_formula('ФОТ_{ОПР без надб.} = p_{ср} N_{пл} m = ' + f'{fn(chapter_2.p_mean)} \\cdot {initial_data.N_pl} \\cdot {len(initial_data.operations)} = {fn(chapter_2.FOT_opr)} [руб./год]')

    dp('Учтём надбавки ОПР:')
    add_formula(
        'ФОТ_{ОПР} = ФОТ_{ОПР} + R_{ОПР} \\cdot (ОПР_{надб.} \\cdot 12 + ОПР_{тариф.ст.}) = ' +
        f'{chapter_2.R_opr} \\cdot ({fn(chapter_2.opr_extra)} \\cdot 12 + {fn(chapter_2.opr_salary)} = {fn(chapter_2.FOT_opr + chapter_2.FOT_opr_extra)} [руб./год]')
    dp('ФОТ ВПР и служащих:')
    p = add_formula_with_description('ФОТ_{ВПР+сл} = \\sum_{i}^{n}{(ТС_i \\cdot N_i \\cdot 12 + ТС_i)} = ' + f'{fn(chapter_2.FOT_vpr + chapter_2.FOT_sl)}', [
        ['n', 'число ОПР и служащих'],
        ['ТС_i', 'тарифная ставка'],
        ['N_i', 'численнойсть']
    ])

    dp('Общий ФОТ:')
    add_formula('ФОТ_{общ} = ФОТ_{ОПР} + ФОТ_{ВПР+сл} = ' + f'{fn(chapter_2.FOT.total)}')

    dp('Таблица 2.3, страховые взносы', table_name_text)
    table = add_table([['Взнос', 'Величина, %', 'Сумма, руб./год']] + [[e.name, fn(e.percent * 100, 1), fn(e.amount)] for e in chapter_2.insurance_fee.rows])
    r = table.add_row()
    r.cells[0].merge(r.cells[1])
    r.cells[0].paragraphs[0].add_run('Итого').bold = True
    r.cells[2].paragraphs[0].add_run(fn(chapter_2.FOT_fee.total)).bold = True
    document.add_page_break()


def gen_1_3():
    dp('3. Формирование сметы затрат на производство', title_text)
    dp('Стоимость основных материалов на единицу изделия:')
    add_formula('S_{ом.ед} = (S_м + S_k) = ' + f'{fn(chapter_3.S_mat_i_comp)} руб./шт.')
    dp('1. Стоимость основных материалов')
    add_formula(
        'S_{ом} = S_{ом.ед} \\cdot N_{пл} = ' + f'{fn(chapter_3.S_mat_i_comp)} \\cdot {fn(initial_data.N_pl, 0)} = {fn(chapter_3.costs["material_main"].total)}  [руб./год]')

    dp(f'2. Стоимость вспомогательных материалов (принято за {fn(chapter_3.help_materials_percent * 100)}% от *)')
    add_formula('S_{вм} = S_{ом} \\cdot k_{вм} = ' + f'{fn(chapter_3.costs["helper"].total)}  [руб./год]')

    dp(f'3. Транспортно-заготовительные расходы (принято за {fn(chapter_3.moving_save_percent * 100)}% от *)')
    add_formula('S_{т-з} = S_{ом} \\cdot k_{т-з} = ' + f'{fn(chapter_3.costs["move save"].total)}  [руб./год]')
    dp(f'Из них {fn(chapter_3.move_save_const_percent * 100, 0)}% - постоянные затраты {fn(chapter_3.costs["move save"].const)} [руб./год]')

    dp(f'4. Инструменты, инвентарь (принято за {fn(chapter_3.inventory_percent * 100)}% от *)')
    add_formula('S_{инстр} = S_{ом} \\cdot k_{интср} = ' + f'{fn(chapter_3.costs["inventory"].total)}  [руб./год]')

    dp(f'5. Топливо и энергия (принято за {fn(chapter_3.fuel_percent * 100)}% от *)')
    add_formula('S_{топл +эн} = S_{ом} \\cdot k_{топл+эн} = ' + f'{fn(chapter_3.costs["fuel total"].total)}  [руб./год]')

    dp(f'5.1 Технологическое ({fn(chapter_3.fuel_tech_percent * 100)}% от топлива и энергии)')
    add_formula('S_{тех. топл + эн} = S_{топл + эн} \\cdot k_{техн} = ' + f'{fn(chapter_3.costs["fuel tech"].total)}  [руб./год]')
    dp(f'5.2 Нетехнологическое ({fn(chapter_3.fuel_non_tech_percent * 100)}% от топлива и энергии)')
    add_formula('S_{тех. топл + эн} = S_{топл + эн} \\cdot (1 - k_{тех.}) = ' + f'{fn(chapter_3.costs["fuel non tech"].total)}  [руб./год]')

    dp('* - от стоимости основных материалов и комплектующих')

    dp(f'Норму амортизации основных средств примем за {fn(chapter_3.OS_amortisation_percent * 100)}%:')
    add_formula('A_{ос} = k_{ам.ос.} \\cdot (S_{осн.ср.} - S_{земл.}) = ' + f'{fn(chapter_3.OS_amortisation)} [руб./год]')

    dp(f'НМА = {fn(chapter_3.NMA)} руб.')
    dp(f'Норма амортизации НМА = {fn(chapter_3.NMA_amortisation_percent * 100)}%:')
    add_formula('A_{НМА} = k_{ам. НМА} \\cdot НМА = ' + f'{fn(chapter_3.NMA_amortisation)} [руб./год]')

    dp(f'Планируемые расходы на ремонт основных средств примем за {fn(chapter_3.OS_fix_percent * 100)}% от стоимости ОС:')
    add_formula('S_{рем.ос} = k_{рем ос} \\cdot (S_{осн.ср.} - S_{земл.}) = ' + f'{fn(chapter_3.OS_fix)} [руб./год]')

    dp('1. Материальные затраты:')
    add_formula('S_{мат.зат} = S_{ом} + S_{вм} + S_{т-з} S_{инстр} S_{топл.+эн.} = ' + f'{fn(chapter_3.costs["material"].total)} [руб./год]')

    dp('2. Затраты на оплату труда:')
    add_formula('S_{ФОТ} = ' + f'{fn(chapter_2.FOT.total)} [руб./год]')

    dp('3. Страховые взносы:')
    add_formula('S_{страх.вз} = ' + f'{fn(chapter_2.FOT_fee.total)} [руб./год]')

    dp('4. Амортизация основных средств и нематериальных активов:')
    add_formula('A_{ОС+НМА} = A_{ос} + A_{НМА}' + f'{fn(chapter_3.costs["amortisation"].total)} [руб./год]')

    dp(f'5. Прочие затраты (примем за {fn(chapter_3.extra_percent * 100)}% от первых 4 пунктов + планируемые расходы на ремонт ОС):')
    add_formula('S_{проч.зат.} = k_{проч.зат.} \\cdot (S_{мат.зат.} + S_{ФОТ} + S_{страх.вз} + A_{ОС+НМА}) + S_{рем.ос} = ' + f'{fn(chapter_3.costs["extra"].total)} [руб./год]')

    costs = [
        chapter_3.costs['material'],
        chapter_3.costs['fot'],
        chapter_3.costs['fot fee'],
        chapter_3.costs['amortisation'],
        chapter_3.costs['extra'],
    ]

    dp('Таблица 3, смета затрат', table_name_text)
    table = add_table(
        [['№', 'Элемент сметы', 'Сумма, руб/год', '%']] +
        [[str(i + 1), e._display_name, fn(e.total), fn(e.total / chapter_3.costs.total * 100)] for i, e in enumerate(costs)],
        [Cm(1), Cm(8), Cm(4), Cm(2)]
    )
    r = table.add_row()
    r.cells[0].merge(r.cells[1])
    r.cells[0].paragraphs[0].add_run('Итого: ').bold = True
    r.cells[0].paragraphs[0].add_run('затраты на производство в текущем периоде ')
    add_formula('S_{пр.тек._{пл.}} ', r.cells[0].paragraphs[0])
    r.cells[2].paragraphs[0].add_run(fn(chapter_3.costs.total)).bold = True
    r.cells[3].paragraphs[0].add_run('100').bold = True
    document.add_page_break()


def gen_1_4():
    dp('4. Расчет себестоимости изделия Б, величины условно-постоянных и переменных затрат', title_text)
    dp('Рассчитаем производственную и полную себестоимости изделия Б')

    dp('4.1 Расчет производственной себестоимости изделия Б', subtitle_text)
    dp('Производственную себестоимость определим по формуле:')
    add_formula('S_{Б\\ произв} = \\frac{S_{пр.тек.пл.}}{N_{пл}} = \\frac{' +
                f'{fn(chapter_3.S_pr_tek_pl)} }}{{ {fn(initial_data.N_pl)} }} = {fn(chapter_4.S_b_proizv)}\\ [руб./шт.]')

    dp('4.2 Расчет полной себестоимости изделия Б', subtitle_text)
    dp('Полную себестоимость определим по формуле:')
    add_formula(
        'S_{Б\\ полн} = \\frac{S_{пр.тек.пл.} + S_{ком}}{N_{пл}} = \\frac{S_{Б сум}}{N_{пл}} = \\frac{' +
        f'{fn(chapter_4.S_sum.total)} }}{{ {fn(initial_data.N_pl)} }} = {fn(chapter_4.S_b_poln.total)}\\ [руб./шт.]')
    dp('Коммерческие затраты связаны с реализацией продукции, включают расходы на тару и упаковку изделий на складах готовой продукции; расходы по '
       'доставке продукции на станцию отправления; комиссионные сборы (отчисления), уплачиваемые сбытовым и другим посредническим предприятиям; '
       'расходы по содержанию помещений для хранения продукции в местах ее реализации; рекламные и представительские расходы)')
    dp(f'Примем коммерческие затраты равными {fn(chapter_4.S_kom_percent * 100)}% от величины затрат на производство в текущем периоде:')
    add_formula('S_{ком} = k_{ком} \\cdot S_{пр.тек.пл.} = ' + f'{fn(chapter_4.S_kom.total)} [руб./год]')
    dp(f'Из них {fn(chapter_4.S_kom_const_percent * 100)}% составляют постоянные затраты, {fn(chapter_4.S_kom.const)}\\ [руб./год]').add_run().add_break(WD_BREAK.PAGE)

    dp('4.3. Построение графических зависимостей (условно-постоянных и переменных затрат)', subtitle_text)

    dp('Таблица 4.3.1, условно-постоянные и переменные затраты', table_name_text)
    table = add_table([
        [f'Суммарные затраты, руб./год: {fn(chapter_4.S_sum.total)}', None, None, None, None, None],
        ['№', 'Условно-постоянные затраты', 'Сумма, тыс.руб./год', '№', 'Переменные затраты', 'Сумма, тыс.руб./год']
    ], [Cm(1), Cm(4.5), Cm(3), Cm(1), Cm(4.5), Cm(3)], style=table_style_12)
    table.cell(0, 0).merge(table.cell(0, 5))
    table.cell(0, 0).paragraphs[0].paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    table.cell(0, 0).paragraphs[0].runs[0].bold = True

    const = [
        (chapter_4.S_kom._display_name, chapter_4.S_kom.const),
        (chapter_3.costs['fuel non tech']._display_name, chapter_3.costs['fuel non tech'].const),
        ('З/п кроме ОПР', chapter_2.FOT.const),
        ('Страховые взносы кроме ОПР', chapter_2.FOT_fee.const),
        (chapter_3.costs['amortisation']._display_name, chapter_3.costs['amortisation'].const),
        (chapter_3.costs['inventory']._display_name, chapter_3.costs['inventory'].const),
        (chapter_3.costs['move save']._display_name, chapter_3.costs['move save'].const),
        (chapter_3.costs['extra']._display_name, chapter_3.costs['extra'].const)
    ]

    variable = [
        (chapter_4.S_kom._display_name, chapter_4.S_kom.variable),
        (chapter_3.costs['fuel tech']._display_name, chapter_3.costs['fuel tech'].variable),
        ('З/п ОПР', chapter_2.FOT.variable),
        ('Страховые взносы ОПР', chapter_2.FOT_fee.variable),
        (chapter_3.costs['material_main']._display_name, chapter_3.costs['material_main'].variable),
        (chapter_3.costs['helper']._display_name, chapter_3.costs['helper'].variable),
        (chapter_3.costs['move save']._display_name, chapter_3.costs['move save'].variable),
        (chapter_3.costs['extra']._display_name, chapter_3.costs['extra'].variable)
    ]

    for i, z in enumerate(zip(const, variable)):
        c = z[0]
        v = z[1]
        r = table.add_row()
        r.cells[0].paragraphs[0].add_run(str(i + 1))
        r.cells[1].paragraphs[0].add_run(c[0])
        r.cells[2].paragraphs[0].add_run(fn(c[1]))
        r.cells[3].paragraphs[0].add_run(str(i + 1))
        r.cells[4].paragraphs[0].add_run(v[0])
        r.cells[5].paragraphs[0].add_run(fn(v[1]))

    r = table.add_row()
    r.cells[0].merge(r.cells[1]).paragraphs[0].add_run('Итого').bold = True
    r.cells[2].paragraphs[0].add_run(fn(sum([e[1] for e in const])))
    r.cells[3].merge(r.cells[4]).paragraphs[0].add_run('Итого').bold = True
    r.cells[5].paragraphs[0].add_run(fn(sum([e[1] for e in variable])))

    add_formula('S_{Б\\ сум} = S_{пост} + S_{перем} = ' + f'{fn(chapter_4.S_sum.total)}')

    ct1 = chapter_4.ct1

    S_const_costs = [e.const for e in ct1.output_data]
    S_variable_costs = [e.variable for e in ct1.output_data]
    S_total_costs = [e.total for e in ct1.output_data]

    B_const_costs = [e.const / i for i, e in ct1.items]
    B_variable_costs = [e.variable / i for i, e in ct1.items]
    B_total_costs = [e.total / i for i, e in ct1.items]

    N_pl_values = chapter_4.N_pl_values
    dp()
    dp('Таблица 4.3.2 зависимость общей суммы затрат на производство и реализацию продукции от величины объема производства за планируемый период', table_name_text)
    add_table([
        ['N, шт./год'] + [fn(e, 0) for e in N_pl_values],
        ['S у-п, руб./год'] + [fn(e) for e in S_const_costs],
        ['S перем, руб./год'] + [fn(e) for e in S_variable_costs],
        ['S сум, руб./год'] + [fn(e) for e in S_total_costs],
        ], style=table_style_10)
    document.add_page_break()

    dp('Таблица 4.3.3, зависимость себестоимости единицы продукции от величины объема производства за планируемый период', table_name_text)
    add_table([
        ['N, шт./год'] + [fn(e, 0) for e in N_pl_values],
        ['B у-п, руб./год'] + [fn(e) for e in B_const_costs],
        ['B перем, руб./год'] + [fn(e) for e in B_variable_costs],
        ['B сум, руб./год'] + [fn(e) for e in B_total_costs],
        ], style=table_style_10)

    dp('График 4.3.4, зависимости общей суммы и себестоимости единицы продукции от величины объёма производства за планируемый период', table_name_text)
    plt.figure(figsize=(10, 4))
    plt.subplot(1, 2, 1)

    plt.title('S(N)')
    plt.xlabel('N, штук')
    plt.ylabel('S, млн. руб./год')
    plt.xticks(N_pl_values, [str(i) for i in N_pl_values], rotation=45)
    plt.yticks(np.linspace(min(S_const_costs + S_variable_costs + S_total_costs) / 1e6, max(S_const_costs + S_variable_costs + S_total_costs) / 1e6, 7))
    plt.grid(True)
    plt.plot(N_pl_values, np.array(S_const_costs) / 1e6, label='S_const')
    plt.plot(N_pl_values, np.array(S_variable_costs) / 1e6, label='S_variable')
    plt.plot(N_pl_values, np.array(S_total_costs) / 1e6, label='S_total')
    plt.legend()

    ax = plt.subplot(1, 2, 2)
    plt.title('B(N)')
    plt.xlabel('N, штук')
    plt.ylabel('B, руб./шт.')
    plt.yscale('log')
    # plt.yticks([0, max(B_const_costs)])
    ticker = matplotlib.ticker.ScalarFormatter()
    ax.yaxis.set_major_formatter(ticker)
    plt.yticks(np.logspace(np.log10(min(B_const_costs + B_variable_costs + B_total_costs)), np.log10(max(B_const_costs + B_variable_costs + B_total_costs)), 7))
    plt.xticks(N_pl_values, [str(i) for i in N_pl_values], rotation=45)
    plt.grid(True)
    plt.plot(N_pl_values, np.array(B_const_costs), label='B_const')
    plt.plot(N_pl_values, np.array(B_variable_costs), label='B_variable')
    plt.plot(N_pl_values, np.array(B_total_costs), label='B_total')
    plt.legend()
    plt.tight_layout()

    memfile = BytesIO()
    plt.savefig(memfile)

    picP = document.add_paragraph()
    picP.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    picP.add_run().add_picture(memfile, width=Cm(17))
    picP.add_run().add_break(WD_BREAK.PAGE)


def gen_1_5():
    dp('5. Расчет потребности в оборотных средствах', title_text)
    dp('Оборотные средства участвуют в одном производственном цикле и полностью переносят свою стоимость на готовую продукцию, и '
       'состоят из оборотных производственных фондов и фондов обращения. Сумма оборотных производственных фондов и фондов обращения '
       'представляют собой оборотные средства предприятия')
    dp()

    dp('5.1 Оборотные средства в производственных запасах', subtitle_text)

    dp('Таблица 5.1.1, норма запаса материалов', table_name_text)
    add_table([['Материал', 'число дней']] + [[e['name'], str(e['t_zap'])] for e in initial_data.materials.rows])

    dp()

    dp('Таблица 5.1.2 норма запаса комплектующих', table_name_text)
    add_table([['Комплектующие', 'число дней']] + [[e['name'], str(e['t_zap'])] for e in initial_data.accessories.rows])

    document.add_page_break()

    dp('Оборотные средства в производственных запасах материалов и комплектующих изделий определим по формуле:')
    add_formula_with_description(
        'K_{об.ср.\\ мат\\ и\\ комп.} = \\sum^{n_{мат}}_{i=1}{\\frac{M_{мат\\ i}N_{пл}}{Т_{пл}}}t_{зап.\\ мат.i} + '
        '\\sum^{n_{комп}}_{j=1}{\\frac{К_{комп\\ j}N_{пл}}{Т_{пл}}}t_{зап.\\ комп.j}', [
            ['М_{мат\\ i},\\ К_{комп\\ j}', 'норма расхода i-го материала и j-го вида комплектующих изделий на одно изготавливаемое изделие в стоимостном выражении, руб./шт.'],
            ['t_{зап.\\ мат.i},\\ t_{зап.\\ комп.j}', 'нормы запаса материалов и комплектующих изделий в календарных днях']
        ]
    )

    add_formula('K_{об.ср.\\ мат\\ и\\ комп.} = ' + f'{fn(chapter_5.K_ob_sr_mk)}\\ [руб.]')

    dp()

    dp(f'Рассчитанную величину оборотных средств в запасах материалов и комплектующих изделий увеличим на {fn(chapter_5.k_ob_sr_percent * 100, 0)}%. '
       f'Общая сумма оборотных средств в производственных запасах составит:')
    add_formula('K_{об.ср._{пр.зап.}} = ' + f'{fn(1 + chapter_5.k_ob_sr_percent)} \\cdot K_{{ об.ср.\\ мат.и\\ комп}} = {fn(chapter_5.K_ob_sr_pr_zap)}\\ [руб.]')
    dp()

    dp('5.2. Оборотные средства в незавершенном производстве', subtitle_text)

    dp('Затраты на материалы и комплектующие изделия:')
    add_formula('S_{мат\\ и\\ комп} =' + f'{fn(chapter_3.S_mat_i_comp)}\\ [руб.]')

    dp('Коэффициент нарастания затрат (условно принимается равномерное нарастание затрат)')
    add_formula(
        'k_{нз} = \\frac{S_{мат\\ и\\ комп} + S_{Б_{произв}}}{2 S_{Б_{произв}}} = \\frac{' +
        f'{fn(chapter_3.S_mat_i_comp)} + {fn(chapter_4.S_b_proizv)} }}{{ 2 \\cdot {fn(chapter_4.S_b_proizv)} }} = {fn(chapter_5.k_nz)}')

    document.add_page_break()

    dp('Производственный цикл (кален. дни) – отрезок времени между началом и окончанием производственного процесса '
       'изготовления одного изделия, включающий время технологических операций; время подготовительно-заключительных операций; '
       'длительность естественных процессов и вспомогательных операций; время межоперационных и междусменных перерывов; '
       'время ожидания обработки при передаче изделий на рабочие места по партиям')

    add_formula_with_description('T_ц = \\frac{\\sum^{m}_{i=1}{t_{техн\\ i}}\\gamma_ц}{C D}\\frac{T_{пл}}{Т_{пл} - B}', [
        ['\\gamma_ц', f'соотношение между производственным циклом и суммарной технологической трудоемкостью изготовления изделия, принято за {fn(chapter_5.gamma_cycle, 0)}']
    ])

    dp('Оборотные средства, находящиеся в незавершенном производстве, руб. рассчитаем по формуле:')
    add_formula('K_{об._{нез.пр.}} = \\frac{S_{Б_{произв}} N_{пл}}{Т_{пл}}k_{нз}T_ц')
    add_formula(f'T_ц = {fn(chapter_5.T_cycle, 3)}\\ [дн.]')
    add_formula(
        'K_{об._{нез.пр.}} = \\frac{' +
        f'{fn(chapter_4.S_b_proizv)} \\cdot {fn(initial_data.N_pl, 0)} }}{{ {chapter_1.T_pl} }} '
        f'\\cdot {fn(chapter_5.k_nz, 2)} \\cdot {fn(chapter_5.T_cycle, 3)} = {fn(chapter_5.K_ob_nez_pr)}\\ [руб.]')

    dp('5.3. Оборотные средства в готовой продукции', subtitle_text)
    dp('Время нахождения на складе:')
    add_formula('t_{реал.} = ' + f'{chapter_5.t_real}\\ [дн.]')
    dp('Оборотные средства, находящиеся в готовой продукции:')
    add_formula(
        'K_{об._{гот.прод.}} = \\frac{S_{Б_{произв}} N_{пл}}{Т_{пл}}t_{реал} = \\frac{' +
        f'{fn(chapter_4.S_b_proizv)} \\cdot {fn(initial_data.N_pl)} }}{{ {chapter_1.T_pl} }} \\cdot {chapter_5.t_real} = {fn(chapter_5.K_ob_got_prod)}\\ [руб.]')

    document.add_page_break()

    dp('5.4. Суммарная потребность в оборотных средствах', subtitle_text)
    dp('Оборотные средства включают в себе не только оборотные средства в производственных запасах, незавершенном производстве и готовой продукции, '
       'а также в расходах будущих периодов, дебиторской задолженности, краткосрочных финансовых вложениях, денежных средствах и т.п. (т.е. прочие оборотные средства).')
    dp('Для упрощения расчетов в курсовой работе зададим удельный вес стоимости производственных запасов, незавершенного производства и готовой продукции в '
       'общей сумме оборотных средств:')

    add_formula('\\gamma_{об} = ' + f'{chapter_5.gamma_ob}')
    dp('Суммарные оборотные средства:')
    add_formula('K_{об_{сум}} = \\frac{K_{об.ср._{пр.зап.}} + K_{об._{нез.пр.}} + K_{об._{гот.прод.}}}{\\gamma_{об}}')
    add_formula('K_{об_{сум}} = \\frac{' + f'{fn(chapter_5.K_ob_sr_pr_zap)} + {fn(chapter_5.K_ob_nez_pr)} + {fn(chapter_5.K_ob_got_prod)} }}'
                                           f'{{ {chapter_5.gamma_ob} }} = {fn(chapter_5.K_ob_sum)}\\ [руб.]')
    dp('Прочие оборотные средства:')
    add_formula('K_{об_{проч}} = (1 - \\gamma_{об}) \\cdot K_{об_{сум}} = ' + f'{fn(chapter_5.K_ob_extra)}\\ [руб.]')

    document.add_page_break()


def gen_1_6():
    dp('6. Бухгалтерский баланс на начало деятельности условного предприятия', title_text)
    dp('Условно в работе вступительный баланс (составляемый на момент возникновения предприятия) совпадает с текущим балансом, составляемым на начало отчетного периода.')
    dp('На начало хозяйственной деятельности в бухгалтерском балансе отсутствуют: затраты в незавершенном производстве; готовая продукция и товары для перепродажи; '
       'дебиторская задолженность. Поэтому их значения на начало деятельности принять равными нулю.')
    dp('Величина сырья и материалов – исходя из рассчитанного норматива оборотных средств в производственных запасов.')
    dp('Значение прочих запасов и затрат определяется как разница между нормативом оборотных средств в производственных запасах с учетом прочих элементов '
       'производственных запасов + РБП, и нормативом оборотных средств в производственных запасов')

    dp('Денежные средства определяются:')
    add_formula(
        'K_{об_{ДС}} = K_{об_{сум}} - (K_{об.ср._{пр.зап.}}) + K_{об._{РБП}}) = ' +
        f'{fn(chapter_5.K_ob_sum)} - ({fn(chapter_5.K_ob_sr_pr_zap)} + {fn(chapter_6.active_passive.K_ob_RBP)}) = {fn(chapter_6.active_passive.K_ob_ds)}\\ [руб.]',
        style=formula_style_12)

    dp()

    dp(f'Удельный вес уставного капитала принят за {fn(chapter_6.ustavnoy_capital_percent * 100)}% от общей суммы пассива баланса.')
    dp('Распределение заёмного капитала:').paragraph_format.first_line_indent = 0
    dp(f'Долгосрочные заёмные средства: {fn(chapter_6.doldosroch_zaemn_sredstva_percent * 100)}%')
    dp(f'Краткосрочные заёмные средства: {fn(chapter_6.kratkosroch_zaemn_sredstva_percent * 100)}%')
    g = 1 - chapter_6.kratkosroch_zaemn_sredstva_percent - chapter_6.doldosroch_zaemn_sredstva_percent
    dp(f'Прочие краткосрочные обязательства: {fn(g * 100)}%').add_run().add_break(WD_BREAK.PAGE)

    dp('Таблица 6, бухгалтерский баланс на начало деятельности условного предприятия', table_name_text)
    add_active_passive_table(chapter_6.active_passive)

    document.add_page_break()


def gen_1_7():
    dp('7. Планирование цены изделия Б, определение фактической ', title_text)
    dp('Расчет планируемой цены изделия Б произведем с помощью методов полных и переменных затрат.')
    dp()

    dp('7. 1 Расчет плановой цены изделия Б методом полных затрат', subtitle_text)
    dp('Расчет осуществим по формуле:')
    add_formula_with_description('Ц_{Б_{полн.(пл)}} = S_{Б\\ полн.} (1 + \\frac{П_{продаж}}{S_{Б\\ сум.}}) = S_{Б\\ полн.}(1 + k_{нац})', [
        ['П_{продаж}', 'прибыль от продаж, условно примем равной прибыли до налогообложения, руб./год'],
        ['k_{нац}', 'коэффициент наценки']
    ])

    dp()
    dp(f'Прибыль до налогообложения определим через удельный вес чистой прибыли в общей сумме прибыли до налогообложения {fn(1 - chapter_7.tax)}')

    dp(f'Условно примем желаемую чистую прибыль за {fn(chapter_7.net_profit_percent * 100)}% от собственного капитала')

    add_formula('П_{чист} = S_{соб.капитал} \\cdot ' + f'{fn(chapter_7.net_profit_percent)} = {fn(chapter_7.net_profit)}\\ [тыс.руб/год]')

    dp('Прибыль от реализации = прибыль до налогообложения:')

    add_formula('П_{реал} = \\frac{П_{чист}}{' + f'{fn(1 - chapter_7.tax)} }} = {fn(chapter_7.profit_before_tax)}\\ [тыс.руб/год]')
    add_formula('k_{нац} = \\frac{' + f'{fn(chapter_7.profit_before_tax)} }}{{ {fn(chapter_4.S_sum.total)} }} = {fn(chapter_7.k_nats)}')
    add_formula(
        'Ц_{Б_{полн.(пл)}} = ' +
        f'{fn(chapter_4.S_b_poln.total)} \\cdot (1 + {fn(chapter_7.k_nats)}) = {fn(chapter_7.P_b_poln)}\\ [руб.]')
    document.add_page_break()

    dp('7.2 Расчет плановой цены изделия Б методом переменных затрат', subtitle_text)
    dp('Расчет осуществим по формуле:')
    add_formula('Ц_{Б_{перем.(пл)}} = S_{Б\\ перем.} (1 + \\frac{П_{продаж} + S_{усл.-пост.}}{S_{перем.}})')
    add_formula(
        'Ц_{Б_{перем.(пл)}} = ' +
        f'{fn(chapter_4.S_b_poln.variable)} (1 + \\frac{{ {fn(chapter_7.profit_before_tax)} + '
        f'{fn(chapter_4.S_sum.const)} }}{{ {fn(chapter_4.S_sum.variable)} }}) = {fn(chapter_7.P_b_perem)}\\ [руб.]')

    dp('Цены должны быть одинаковыми, если не округлять до копеек дельные затраты и коэффициент наценки. При вычислении использованы более точные цифры, '
       'чем те, которые указаны в уравнении, поэтому разница минимальна, или отстутсвует вовсе.')
    dp()

    dp(f'В качестве плановой цены возьмём большую из двух: {fn(chapter_7.P_proizv_plan)} [руб.]')
    dp()

    dp('7.3. Расчет рыночной (фактической) цены изделия Б', subtitle_text)
    dp('Реальные условия сложились таким образом, что изделие Б смогли реализовать по цене ниже, чем запланировали:')
    add_formula(
        'Ц_{Б\\ факт} = Ц_{Б\\ произв\\ пл} \\cdot k_{Ц\\ факт} = ' +
        f' {fn(chapter_7.P_proizv_plan)} \\cdot {fn(chapter_7.price_fact_percent)} = {fn(chapter_7.P_fact)}\\ [руб.]')

    document.add_page_break()


def gen_1_8():
    dp('8. Отчет о финансовых результатах на конец первого периода', title_text)
    dp(f'Примем что реализовать удалось только {fn(chapter_8.N_fact_percent * 100)}% продукции:')
    add_formula('N_{факт} = ' + f'{fn(chapter_8.N_fact_percent)} \\cdot N_{{пл}} = {chapter_8.N_fact}\\ [шт.]')

    dp('Определим выручку:')
    add_formula('Q_{план} = Ц_{Б\\ произв\\ пл} \\cdot N_{пл} = ' + f'{fn(chapter_8.P_pr_plan)} \\cdot {fn(initial_data.N_pl)} = {fn(chapter_8.Q_plan)}')
    add_formula('Q_{факт} = Ц_{Б\\ произв\\ факт} \\cdot N_{факт} = ' + f'{fn(chapter_8.P_pr_fact)} \\cdot {fn(chapter_8.N_fact)} = {fn(chapter_8.Q_fact)}')
    dp()

    dp('Определим себестоимость проданной готовой продукции:')
    add_formula(
        'S_{пр.гот.пр_{план}} = S_{пр.тек_{пл.}} - K_{об_{нез.пр.}} - K_{об._{гот.прод.}} = ' +
        f'{fn(chapter_3.S_pr_tek_pl)} - {fn(chapter_5.K_ob_nez_pr)} - {fn(chapter_5.K_ob_got_prod)} = {fn(chapter_8.S_pr_got_pr_plan)}', style=formula_style_12)
    add_formula(
        'S_{пр.гот.пр_{факт}} = S_{пр.тек_{пл.}} - K_{об_{нез.пр.}} - K_{об._{гот.прод._{факт}}} = '
        'S_{пр.тек_{пл.}} - K_{об_{нез.пр.}} - S_{Б_{произв}} N_{ост.} - K_{об._{гот.прод.}} = ' +
        f'{fn(chapter_3.S_pr_tek_pl)} - {fn(chapter_5.K_ob_nez_pr)} - ({fn(chapter_4.S_b_poln.total)}'
        f' \\cdot {fn(chapter_8.N_ost, 0)}) = {fn(chapter_8.S_pr_got_pr_fact)}', style=formula_style_12)

    dp()
    dp('Определим валовую прибыль:')
    add_formula(
        'S_{валовая\\ план} = Q_{план} - S_{пр.гот.пр_{план}} = ' +
        f'{fn(chapter_8.Q_plan)} - {fn(chapter_8.S_pr_got_pr_plan)} = {fn(chapter_8.S_valovaya_plan)}', style=formula_style_12)
    add_formula(
        'S_{валовая\\ факт} = Q_{факт} - S_{пр.гот.пр_{факт}} = ' +
        f'{fn(chapter_8.Q_fact)} - {fn(chapter_8.S_pr_got_pr_fact)} = {fn(chapter_8.S_valovaya_fact)}', style=formula_style_12)

    dp()
    dp(f'Фактические коммерческие расходы определим как {fn(chapter_8.kom_percent * 100, 0)}% от планируемых')
    add_formula('K_{ком\\ план} = ' + f'{fn(chapter_8.S_kom_plan)}\\ (см\\ п.\\ 4.2)')
    add_formula('K_{ком\\ факт} = K_{ком\\ план} \\cdot ' + f'{fn(chapter_8.kom_percent)} = {fn(chapter_8.S_kom_fact)}')

    document.add_page_break()
    dp('Определим прибыль от продаж:')
    add_formula('П_{пр_{план}} = S_{валовая\\ план} - K_{ком.пл.} = ' +
                f'{fn(chapter_8.S_valovaya_plan)} - {fn(chapter_8.S_kom_plan)} = {fn(chapter_8.P_pr_plan)}', style=formula_style_12)
    add_formula('П_{пр_{факт}} = S_{валовая\\ факт} - K_{ком.факт.} = ' +
                f'{fn(chapter_8.S_valovaya_fact)} - {fn(chapter_8.S_kom_fact)} = {fn(chapter_8.P_pr_fact)}', style=formula_style_12)

    dp('Определим прочие расходы (прочие доходы примем равными нулю):')
    dp(f'Для этого условно вычтем из прибыли от продаж прибыль до налогообложения (см. п. 7.1), '
       f'фактические прочие расходы определим как {fn(chapter_8.pr_dir_fact_percent * 100, 0)}% от плановых.')
    add_formula('S_{прочие\\ план.} = П_{пр_{план}} - П_{продаж.} = ' +
                f'{fn(chapter_8.P_pr_plan)} - {fn(chapter_8.P_pr_do_nalogov_plan)} = {fn(chapter_8.S_prochie_dohidy_i_rashody_plan)}', style=formula_style_12)
    add_formula('S_{прочие\\ факт.} = S_{прочие план.} \\cdot ' +
                f'{fn(chapter_8.pr_dir_fact_percent)}  = {fn(chapter_8.S_prochie_dohidy_i_rashody_fact)}', style=formula_style_12)

    dp()
    dp('Определим прибыль до налогообложения:')
    add_formula(
        'П_{до\\ налогообл.\\ план.} = П_{пр_{план}} - S_{прочие\\ план.} = ' +
        f'{fn(chapter_8.P_pr_plan)} - {fn(chapter_8.S_prochie_dohidy_i_rashody_plan)} = {fn(chapter_8.P_pr_do_nalogov_plan)}', style=formula_style_12)
    add_formula(
        'П_{до\\ налогообл.\\ факт.} = П_{пр_{факт}} - S_{прочие\\ факт.} = ' +
        f'{fn(chapter_8.P_pr_fact)} - {fn(chapter_8.S_prochie_dohidy_i_rashody_fact)} = {fn(chapter_8.P_pr_do_nalogov_fact)}', style=formula_style_12)

    dp()
    dp(f'Определим налог на прибыль (для ООО по общей системе налогообложения составляет {fn(chapter_7.tax * 100)}%):')
    add_formula('S_{налог\\ план.} = П_{до\\ налогообл.\\ план.} \\cdot ' + f'{fn(chapter_7.tax)} = {fn(chapter_8.nalog_na_pribil_plan)}')
    add_formula('S_{налог\\ факт.} = П_{до\\ налогообл.\\ факт.} \\cdot ' + f'{fn(chapter_7.tax)} = {fn(chapter_8.nalog_na_pribil_fact)}')

    dp()
    dp('Определим чистую прибыль:')
    add_formula(
        'П_{чист.\\ план.} = П_{до\\ налогообл.\\ план.} - S_{налог\\ план.} = ' +
        f'{fn(chapter_8.P_pr_do_nalogov_plan)} - {fn(chapter_8.nalog_na_pribil_plan)} = {fn(chapter_8.P_chistaya_plan)}', style=formula_style_12)
    add_formula(
        'П_{чист.\\ факт.} = П_{до\\ налогообл.\\ факт.} - S_{налог\\ факт.} = ' +
        f'{fn(chapter_8.P_pr_do_nalogov_fact)} - {fn(chapter_8.nalog_na_pribil_fact)} = {fn(chapter_8.P_chistaya_fact)}', style=formula_style_12)
    document.add_page_break()

    dp('Таблица 8, отчёт о финансовых результатах на конец первого периода', table_name_text)
    table = add_table([
        ['Наименование показателя', 'Сумма, руб/год', None],
        [None, 'план', 'факт'],
        ['Выручка', fn(chapter_8.Q_plan), fn(chapter_8.Q_fact)],
        ['Себестоимость продаж (проданной готовой продукции)', fn(chapter_8.S_pr_got_pr_plan), fn(chapter_8.S_pr_got_pr_fact)],
        ['Валовая прибыль (убыток)', fn(chapter_8.S_valovaya_plan), fn(chapter_8.S_pr_got_pr_fact)],
        ['Коммерческие расходы', fn(chapter_8.S_kom_plan), fn(chapter_8.S_kom_plan)],
        ['Прибыль (убыток) от продаж', fn(chapter_8.P_pr_plan), fn(chapter_8.P_pr_fact)],
        ['Прочие доходы ', fn(0), fn(0)],
        ['Прочие расходы', fn(chapter_8.S_prochie_dohidy_i_rashody_plan), fn(chapter_8.S_prochie_dohidy_i_rashody_fact)],
        ['Прибыль (убыток) до налогообложения', fn(chapter_8.P_pr_do_nalogov_plan), fn(chapter_8.P_pr_do_nalogov_fact)],
        ['Налог на прибыль', fn(chapter_8.nalog_na_pribil_plan), fn(chapter_8.nalog_na_pribil_fact)],
        ['Чистая прибыль (убыток)', fn(chapter_8.P_chistaya_plan), fn(chapter_8.P_chistaya_fact)],

    ], [Cm(11), Cm(6.25), None])
    table.cell(0, 0).merge(table.cell(1, 0))
    table.cell(0, 1).merge(table.cell(0, 2))
    document.add_page_break()


def gen_1_9():
    dp('9. Плановый и фактический бухгалтерский баланс на конец периода', title_text)

    dp(f'Амортизация основных средств: {fn(chapter_3.costs["amortisation OS"].total)} [руб./год] (см п.3)')
    dp(f'Амортизация НМА: {fn(chapter_3.costs["amortisation NMA"].total)} [руб./год] (см п.3)')
    dp('Рассчитаем оборотные средства в незавершённом производстве и готовой продукции:')
    add_formula('K_{об_{нез.пр.\\ план}} = ' + f'{fn(chapter_5.K_ob_nez_pr)}\\ [руб.]\\ (см.\\ п.\\ 5.2)')
    add_formula('K_{об_{гот.пр.\\ план}} = ' + f'{fn(chapter_5.K_ob_got_prod)}\\ [руб.]\\ (см.\\ п.\\ 5.2)')
    add_formula('K_{об_{гот.под.\\ факт}} = K_{об._{гот.прод.}} + S_{Б_{произв}} N_{ост.} = ' + f'{fn(chapter_8.K_ob_got_prod_fact)}\\ [руб.]')

    dp()
    dp('Рассчитаем денежные средства:')
    add_formula('K_{ден.ср.\\ план} = K_{ден.ср.нач.период.} - A_{НМА} - А_{ОС} + П_{чист.\\ план} - (K_{об_{нез.пр.\\ план}} + K_{об_{гот.пр.\\ план}}) = ' +
                f'{fn(chapter_6.active_passive.K_ob_ds)} - {fn(chapter_3.costs["amortisation NMA"].total)} - {fn(chapter_3.costs["amortisation OS"].total)} + '
                f'{fn(chapter_8.P_chistaya_plan)} - ({fn(chapter_5.K_ob_nez_pr)} + {fn(chapter_8.K_ob_got_prod_plan)}) = {fn(chapter_9.K_den_sr_plan)}', style=formula_style_12)
    add_formula('K_{ден.ср.\\ факт} = K_{ден.ср.нач.период.} - A_{НМА} - А_{ОС} + П_{чист.\\ факт} - (K_{об_{нез.пр.\\ факт}} + K_{об_{гот.пр.\\ факт}}) = ' +
                f'{fn(chapter_6.active_passive.K_ob_ds)} - {fn(chapter_3.costs["amortisation NMA"].total)} - {fn(chapter_3.costs["amortisation OS"].total)} + '
                f'{fn(chapter_8.P_chistaya_fact)} - ({fn(chapter_5.K_ob_nez_pr)} + {fn(chapter_8.K_ob_got_prod_fact)}) = {fn(chapter_9.K_den_sr_fact)}', style=formula_style_12)

    dp()
    if chapter_9.valid_to_cope_kz_plan == 'full':
        dp('Возможно полное погашение краткосрочных заёмных средств для плановой суммы денежных средств:')
        add_formula('K_{ден.ср.конец\\ план} = K_{ден.ср.\\ план} - S_{кр.заёмн.ср.} = ' +
                    f'{fn(chapter_9.K_den_sr_plan)} - {fn(chapter_6.active_passive.kratkosroch_zaem_sredstva)} = {fn(chapter_9.K_den_sr_konez_plan)}\\ [руб.]',
                    style=formula_style_12)
    elif chapter_9.valid_to_cope_kz_plan == 'part':
        dp('Возможно частичное погашение краткосрочных заёмных средств для плановой суммы '
           'денежных средств (дальнейшая генерация неверна):').runs[0].font.color.rgb = RGBColor(255, 0, 0)
        add_formula('K_{ден.ср.конец\\ план} = K_{ден.ср.\\ план} - S_{кр.заёмн.ср.} = ' + f'{fn(500_000)} \\ [руб.]')
        add_formula('S_{кр.заёмн.ср.\\ план} = ' + f'{fn(chapter_9.S_kratkosroch_zaem_sredstva_konez_plan)} \\ [руб.]')
    else:
        dp('Погашение краткосрочных заёмных средств невозможно для плановой суммы денежных средств, '
           'дальнейшая генерация неверна.').runs[0].font.color.rgb = RGBColor(255, 0, 0)

    if chapter_9.valid_to_cope_kz_fact == 'full':
        dp('Возможно полное погашение краткосрочных заёмных средств для фактической суммы денежных средств:')
        add_formula('K_{ден.ср.конец\\ факт} = K_{ден.ср.\\ факт} - S_{кр.заёмн.ср.} = ' +
                    f'{fn(chapter_9.K_den_sr_fact)} - {fn(chapter_6.active_passive.kratkosroch_zaem_sredstva)} = {fn(chapter_9.K_den_sr_konez_fact)}\\ [руб.]',
                    style=formula_style_12)
    elif chapter_9.valid_to_cope_kz_fact == 'part':
        dp('Возможно частичное погашение краткосрочных заёмных средств для фактической суммы '
           'денежных средств (дальнейшая генерация неверна):').runs[0].font.color.rgb = RGBColor(255, 0, 0)
        add_formula('K_{ден.ср.конец\\ факт} = K_{ден.ср.\\ факт} - S_{кр.заёмн.ср.} = ' + f'{fn(500_000)} \\ [руб.]')
        add_formula('S_{кр.заёмн.ср.\\ факт} = ' + f'{fn(chapter_9.S_kratkosroch_zaem_sredstva_konez_fact)} \\ [руб.]')
    else:
        dp('Погашение краткосрочных заёмных средств невозможно для фактической суммы денежных средств, '
           'дальнейшая генерация неверна.').runs[0].font.color.rgb = RGBColor(255, 0, 0)

    document.add_page_break()

    dp('Таблица 9.1, плановый бухгалтерский баланс на конец периода', table_name_text)
    add_active_passive_table(chapter_9.active_passive_plan)

    document.add_page_break()

    dp('Таблица 9.2, фактический бухгалтерский баланс на конец периода', table_name_text)
    add_active_passive_table(chapter_9.active_passive_fact)
    document.add_page_break()


def gen_1_10():
    dp('10. Анализ деятельности условного предприятия', title_text)
    dp('Анализ деятельности условного предприятия осуществляется на основе основных плановых и фактических '
       'показателей хозяйственной деятельности предприятия и графика рентабельности.')

    def gen_1_10_1():
        dp()
        dp('10.1 Основные показатели хозяйственной деятельности предприятия', subtitle_text)

        dp()
        dp('Сумма хозяйственных средств:')
        add_formula('K_{хс.\\ план} = ' + f'{fn(chapter_9.active_passive_plan.active)}\\ [руб.]', style=formula_style_12)
        add_formula('K_{хс.\\ факт} = ' + f'{fn(chapter_9.active_passive_fact.active)}\\ [руб.]', style=formula_style_12)

        dp()
        dp('Собственные оборотные средства:')
        add_formula('k_{соб.об.ср.} = Оборотные\\ активы - Краткосрочные\\ обязательства')
        add_formula('k_{соб.об.ср.\\ план} = ' +
                    f'{fn(chapter_9.active_passive_plan.r2)} - {fn(chapter_9.active_passive_plan.r5)} = {fn(chapter_10.k_sob_ob_sr_plan)}\\ [руб.]', style=formula_style_12)
        add_formula('k_{соб.об.ср.\\ факт} = ' +
                    f'{fn(chapter_9.active_passive_fact.r2)} - {fn(chapter_9.active_passive_fact.r5)} = {fn(chapter_10.k_sob_ob_sr_fact)}\\ [руб.]', style=formula_style_12)
        dp()
        dp('Коэффициент обеспеченности собственными средствами:')
        add_formula('k_{обеспеч.соб.ср.} = \\frac{Оборотные\\ активы - Краткосрочные\\ обязательства}{Оборотные\\ активы}')
        add_formula(
            'k_{обеспеч.соб.ср.\\ план} = \\frac{' +
            f'{fn(chapter_10.k_sob_ob_sr_plan)} }}{{ {fn(chapter_9.active_passive_plan.r2)} }} = {fn(chapter_10.k_obespech_sob_sr_plan)}', style=formula_style_12)
        add_formula(
            'k_{обеспеч.соб.ср.\\ факт} = \\frac{' +
            f'{fn(chapter_10.k_sob_ob_sr_fact)}  }}{{ {fn(chapter_9.active_passive_fact.r2)} }} = {fn(chapter_10.k_obespech_sob_sr_fact)}', style=formula_style_12)

        document.add_page_break()

        dp('Коэффициент абсолютной ликвидности:')
        add_formula('k_{абс.ликв.} = \\frac{Абсолютно\\ ликвидныке\\ активы}{Краткосрочные\\ обязательства}')
        add_formula('k_{абс.ликв.\\ план} = \\frac{' +
                    f'{fn(chapter_9.active_passive_plan.K_ob_ds)} }}{{ {fn(chapter_9.active_passive_plan.r5)} }} = {fn(chapter_10.k_abs_likvid_plan)}', style=formula_style_12)
        add_formula('k_{абс.ликв.\\ факт} = \\frac{' +
                    f'{fn(chapter_9.active_passive_fact.K_ob_ds)} }}{{ {fn(chapter_9.active_passive_fact.r5)} }} = {fn(chapter_10.k_abs_likvid_fact)}', style=formula_style_12)

        dp()
        dp('Коэффициент текущей ликвидности (или коэффициент покрытия баланса):')
        add_formula('k_{тек.ликв.} = \\frac{Сумма\\ оборотных\\ активов}{Краткосрочные\\ обязательства}')
        add_formula('k_{тек.ликв.\\ план} = \\frac{' +
                    f'{fn(chapter_9.active_passive_plan.r2)} }}{{ {fn(chapter_9.active_passive_plan.r5)} }} = {fn(chapter_10.k_tek_likvid_plan)}', style=formula_style_12)
        add_formula('k_{тек.ликв.\\ факт} = \\frac{' +
                    f'{fn(chapter_9.active_passive_fact.r2)} }}{{ {fn(chapter_9.active_passive_fact.r5)} }} = {fn(chapter_10.k_tek_likvid_fact)}', style=formula_style_12)

        dp()
        dp('Выручка от продажи продукции:')
        add_formula('Q_{план} = ' + f'{fn(chapter_8.Q_plan)}\\ [руб.]', style=formula_style_12)
        add_formula('Q_{факт} = ' + f'{fn(chapter_8.Q_fact)}\\ [руб.]', style=formula_style_12)

        dp()
        dp('Нераспределенная прибыль:')
        add_formula('П_{нерасп.\\ план} = ' + f'{fn(chapter_9.active_passive_plan.neraspred_pribil)}\\ [руб.]', style=formula_style_12)
        add_formula('П_{нерасп.\\ факт} = ' + f'{fn(chapter_9.active_passive_fact.neraspred_pribil)}\\ [руб.]', style=formula_style_12)

        document.add_page_break()

        dp('Выработка продукции на одного работника:')
        add_formula('V = \\frac{Объём\\ продукции}{Среднесписочное\\ кол-во\\ ППП} = \\frac{' +
                    f'{fn(initial_data.N_pl, 0)} }}{{ {fn(chapter_2.R_ppp, 0)} }} = {fn(chapter_10.V)}\\ [шт./работн.год]', style=formula_style_12)

        dp()
        dp('Среднегодовая стоимость ОПФ:')
        add_formula(
            'S_{ср.год.ст.ОПФ} = S_{ОПФ\\ нач.пер.} - А_{ОПФ} \\cdot 0.5 = ' +
            f'{fn(chapter_1.S_os_amortisable)} - {fn(chapter_3.costs["amortisation OS"].total)} = {fn(chapter_10.OS_year_mean)}\\ [руб.]', style=formula_style_12)

        dp()
        dp('Коэффициент фондоотдачи:')
        add_formula('k_{ФО\\ план} = \\frac{Q_{план}}{Среднегодовая\\ стоимость\\ ОПФ} = \\frac{' +
                    f'{fn(chapter_8.Q_plan)} }}{{ {fn(chapter_10.OS_year_mean)} }} = {fn(chapter_10.k_FO_plan)}', style=formula_style_12)
        add_formula('k_{ФО\\ факт} = \\frac{Q_{факт}}{Среднегодовая\\ стоимость\\ ОПФ} = \\frac{' +
                    f'{fn(chapter_8.Q_fact)} }}{{ {fn(chapter_10.OS_year_mean)} }} = {fn(chapter_10.k_FO_fact)}', style=formula_style_12)

        dp()
        dp('Коэффициент фондоемкости:')
        add_formula('k_{ФЕ\\ план} = k_{ФО\\ план}^{-1} = ' + f'{fn(chapter_10.k_FO_plan)} ^ {{-1}} = {fn(chapter_10.k_FE_plan)}', style=formula_style_12)
        add_formula('k_{ФЕ\\ факт} = k_{ФО\\ факт}^{-1} = ' + f'{fn(chapter_10.k_FO_fact)} ^ {{-1}} = {fn(chapter_10.k_FE_fact)}', style=formula_style_12)

        dp()
        dp('Число оборотов оборотных средств:')
        add_formula('Ср.сумм.исп.об.ср._{план} = ' + f'{fn(chapter_10.Z_ob_sr_year_mean_plan)}\\ [руб.]', style=formula_style_12)
        add_formula('Ср.сумм.исп.об.ср._{факт} = ' + f'{fn(chapter_10.Z_ob_sr_year_mean_fact)}\\ [руб.]', style=formula_style_12)
        add_formula('Z_{об} = \\frac{Выручка\\ от\\ реализации}{Средняя\\ сумма\\ используемых\\ обороных\\ средств}')
        add_formula('Z_{об\\ план} = \\frac{' +
                    f'{fn(chapter_8.Q_plan)} }}{{ {fn(chapter_10.Z_ob_sr_year_mean_plan)} }} = {fn(chapter_10.Z_ob_sr_year_mean_plan)} [раз/год]', style=formula_style_12)
        add_formula('Z_{об\\ факт} = \\frac{' +
                    f'{fn(chapter_8.Q_fact)} }}{{ {fn(chapter_10.Z_ob_sr_year_mean_fact)} }} = {fn(chapter_10.Z_ob_sr_year_mean_fact)} [раз/год]', style=formula_style_12)

        document.add_page_break()

        dp('Оборачиваемость собственного капитала:')
        add_formula('Ср.год.собств.кап_{план} = ' + f'{fn(chapter_10.S_sobstv_cap_year_mean_plan)}\\ [руб.]', style=formula_style_12)
        add_formula('Ср.год.собств.кап._{факт} = ' + f'{fn(chapter_10.S_sobstv_cap_year_mean_fact)}\\ [руб.]', style=formula_style_12)
        add_formula('k_{об.собств.кап.} = \\frac{Выручка\\ от\\ реализации}{Ср.год.собств.кап}')
        add_formula('k_{об.собств.кап.\\ план} = \\frac{' +
                    f'{fn(chapter_8.Q_plan)} }}{{ {fn(chapter_10.S_sobstv_cap_year_mean_plan)} }} = {fn(chapter_10.k_oborach_sobstv_capital_plan)}', style=formula_style_12)
        add_formula('k_{об.собств.кап.\\ факт} = \\frac{' +
                    f'{fn(chapter_8.Q_fact)} }}{{ {fn(chapter_10.S_sobstv_cap_year_mean_fact)} }} = {fn(chapter_10.k_oborach_sobstv_capital_fact)}', style=formula_style_12)

        dp()
        dp('Рентабельность продукции:')
        add_formula('R_{продукции} = \\frac{Прибыль\\ от\\ продаж}{Себестоимость\\ продаж}')
        add_formula('R_{продукции\\ план} = \\frac{' +
                    f'{fn(chapter_8.P_pr_plan)} }}{{ {fn(chapter_4.S_sum.total)} }} = {fn(chapter_10.R_production_plan)}', style=formula_style_12)
        add_formula('R_{продукции\\ факт} = \\frac{' +
                    f'{fn(chapter_8.P_pr_fact)} }}{{ {fn(chapter_4.S_sum.total)} }} = {fn(chapter_10.R_production_fact)}', style=formula_style_12)

        dp()
        dp('Рентабельность продаж:')
        add_formula('R_{продаж} = \\frac{Чистая\\ прибыль}{Выручка}')
        add_formula('R_{продаж\\ план} = \\frac{' +
                    f'{fn(chapter_8.P_chistaya_plan)} }}{{ {fn(chapter_8.Q_plan)} }} = {fn(chapter_10.R_sell_plan)}', style=formula_style_12)
        add_formula('R_{продаж\\ факт} = \\frac{' +
                    f'{fn(chapter_8.P_chistaya_fact)} }}{{ {fn(chapter_8.Q_fact)} }} = {fn(chapter_10.R_sell_fact)}', style=formula_style_12)

        document.add_page_break()
        dp('Рентабельность активов:')
        add_formula('R_{активов} = \\frac{Чистая\\ прибыль}{Актив}')
        add_formula('R_{активов\\ план} = \\frac{' +
                    f'{fn(chapter_8.P_chistaya_plan)} }}{{ {fn(chapter_9.active_passive_plan.active)} }} = {fn(chapter_10.R_active_plan)}', style=formula_style_12)
        add_formula('R_{активов\\ факт} = \\frac{' +
                    f'{fn(chapter_8.P_chistaya_fact)} }}{{ {fn(chapter_9.active_passive_fact.active)} }} = {fn(chapter_10.R_active_fact)}', style=formula_style_12)

        dp()
        dp('Рентабельность собственного капитала:')
        add_formula('R_{собств.кап.} = \\frac{Чистая\\ прибыль}{Актив}')
        add_formula('R_{собств.кап.\\ план} = \\frac{' +
                    f'{fn(chapter_8.P_chistaya_plan)} }}{{ {fn(chapter_10.S_sobstv_cap_year_mean_plan)} }} = {fn(chapter_10.R_sobstv_capital_plan)}', style=formula_style_12)
        add_formula('R_{собств.кап.\\ факт} = \\frac{' +
                    f'{fn(chapter_8.P_chistaya_fact)} }}{{ {fn(chapter_10.S_sobstv_cap_year_mean_plan)} }} = {fn(chapter_10.R_sobstv_capital_fact)}', style=formula_style_12)

        document.add_page_break()

        dp('Таблица 10, Плановые и фактические значения основных показателей хозяйственной деятельности предприятия в I периоде', table_name_text)
        add_table([
            ['Наименование показателя и его размерность', 'План', 'Факт'],
            ['Сумма хозяйственных средств, [руб.]', fn(chapter_9.active_passive_plan.active), fn(chapter_9.active_passive_fact.active)],
            ['Собственные оборотные средства, [руб.]', fn(chapter_10.k_sob_ob_sr_plan), fn(chapter_10.k_sob_ob_sr_fact)],
            ['Коэффициент обеспеченности собственными средствами', fn(chapter_10.k_obespech_sob_sr_plan), fn(chapter_10.k_obespech_sob_sr_fact)],
            ['Коэффициент абсолютной ликвидности', fn(chapter_10.k_abs_likvid_plan), fn(chapter_10.k_abs_likvid_fact)],
            ['Коэффициент текущей ликвидности', fn(chapter_10.k_tek_likvid_plan), fn(chapter_10.k_tek_likvid_fact)],
            ['Выручка от продажи продукции, [руб.]', fn(chapter_8.Q_plan), fn(chapter_8.Q_fact)],
            ['Нераспределенная прибыль, [руб.]', fn(chapter_9.active_passive_plan.neraspred_pribil), fn(chapter_9.active_passive_fact.neraspred_pribil)],
            ['Выработка продукции на одного работника [шт./работн.год]', fn(chapter_10.V), fn(chapter_10.V)],
            ['Среднегодовая стоимость ОПФ, [руб.]', fn(chapter_10.OS_year_mean), fn(chapter_10.OS_year_mean)],
            ['Коэффициент фондоотдачи', fn(chapter_10.k_FO_plan), fn(chapter_10.k_FO_fact)],
            ['Коэффициент фондоемкости', fn(chapter_10.k_FE_plan), fn(chapter_10.k_FE_fact)],
            ['Число оборотов оборотных средств, [раз/год]', fn(chapter_10.Z_ob_sr_year_mean_plan), fn(chapter_10.Z_ob_sr_year_mean_fact)],
            ['Оборачиваемость собственного капитала', fn(chapter_10.k_oborach_sobstv_capital_plan), fn(chapter_10.k_oborach_sobstv_capital_fact)],
            ['Рентабельность продукции', fn(chapter_10.R_production_plan), fn(chapter_10.R_production_fact)],
            ['Рентабельность продаж', fn(chapter_10.R_sell_plan), fn(chapter_10.R_sell_fact)],
            ['Рентабельность активов', fn(chapter_10.R_active_plan), fn(chapter_10.R_active_fact)],
            ['Рентабельность собственного капитала', fn(chapter_10.R_sobstv_capital_plan), fn(chapter_10.R_sobstv_capital_fact)],
        ], [Cm(9.25), Cm(3.75), Cm(3.75)], True, style=table_style_12)

        document.add_page_break()

    gen_1_10_1()
    dp('10.2 График рентабельности изделия Б', subtitle_text)
    dp('Построим график рентабельности в соответствии с полученными значениями, рассчитаем точку безубыточности, '
       'коэффициент покрытия, запас финансовой прочности и величину операционного рычага для плановых условий.')

    dp('Точку безубыточности (критический объем продаж) определим:')
    add_formula('N_{кр} = \\frac{S_{усл.пост.}}{Ц_{Б\\ произв\\ план} - S_{Б\\ перем}}')
    p = dp('Замечу, что предоставленная формула не учитывает того, что ')
    add_formula('S_{Б\\ перем}}', p)
    p.add_run(' может меняться от объёма продукции (см п. 4.3), поэтому воспользуемся иной формулой:')
    add_formula('N_{кр} \\cdot {Ц_{Б\\ произв\\ план} = S_{сум}(N_{кр})')
    add_formula('Q_{кр} = N_{кр} \\cdot {Ц_{Б\\ произв\\ план}} = ' + f'{fn(chapter_10.N_kr, 0)} \\cdot {chapter_7.P_proizv_plan} = {fn(chapter_10.Q_kr)}')
    dp('Решение найдём итеративно, с помощью бинарного поиска:')
    add_formula('N_{кр} = ' + f'{fn(chapter_10.N_kr, 0)}')

    document.add_page_break()
    dp('Коэффициент покрытия:')
    add_formula('k_{покр} = \\frac{Ц_{Б\\ произв\\ план} - S_{Б\\ перем}}{Ц_{Б\\ произв\\ план}')
    p = dp('Здесь формула снова не учитывает того, что ')
    add_formula('S_{Б\\ перем}}', p)
    p.add_run(' может меняться от объёма продукции:')
    add_formula('k_{покр\\ N} = \\frac{Ц_{Б\\ произв\\ план} - S_{Б\\ перем}(N)}{Ц_{Б\\ произв\\ план}')

    table = add_table([['N'] + [fn(e, 0) for e in chapter_10.k_pokr.input_data], [''] + [fn(e, 3) for e in chapter_10.k_pokr.output_data]])
    add_formula('k_{покр}', table.cell(1, 0).paragraphs[0])

    dp()
    dp('Запас финансовой прочности:')
    add_formula('Q_{фин\\ пр.} = \\frac{Q - Q_{кр}}{Q} \\cdot 100%')
    add_formula('Q_{фин\\ пр.\\ план} = \\frac{ ' + f'{fn(chapter_8.Q_plan)} - {chapter_10.Q_kr} }}{{ {fn(chapter_8.Q_plan)} }} = {fn(chapter_10.Q_fin_pr_plan)}')
    add_formula('Q_{фин\\ пр.\\ факт} = \\frac{ ' + f'{fn(chapter_8.Q_fact)} - {chapter_10.Q_kr} }}{{ {fn(chapter_8.Q_fact)} }} = {fn(chapter_10.Q_fin_pr_fact)}')

    dp()
    dp('Найдём эффект производственного рычага:')
    add_formula('E_{пр.\\ рыч.} = \\frac{Маржинальная\\ прибыль}{Прибыль\\ от\\ продаж}')
    add_formula(
        'E_{пр.\\ рыч.\\ план} = \\frac{ ' +
        f'{fn(chapter_8.Q_plan)} - {chapter_4.S_sum.total} }}{{ {fn(chapter_8.P_pr_plan)} }} = {fn(chapter_10.proizv_richag_plan)}')
    add_formula(
        'E_{пр.\\ рыч.\\ факт} = \\frac{ ' +
        f'{fn(chapter_8.Q_plan)} - {fn(chapter_4.S_b_poln.variable)} \\cdot {fn(chapter_8.N_fact)} }}{{ {fn(chapter_8.P_pr_fact)} }} = {fn(chapter_10.proizv_richag_fact)}')

    document.add_page_break()

    dp('График 10, рентабельность изделия Б')

    plt.figure(figsize=(8, 8))
    plt.subplot(1, 1, 1)

    ct1 = chapter_4.ct1

    S_const_costs = [e.const for e in ct1.output_data]
    S_variable_costs = [e.variable for e in ct1.output_data]
    S_total_costs = [e.total for e in ct1.output_data]

    N_pl_values = chapter_4.N_pl_values

    plt.title('S(N)')
    plt.xlabel('N, шт. / год')
    plt.ylabel('Выручка, затраты, тыс. руб./год')
    plt.xticks([0, chapter_10.N_kr, initial_data.N_pl], ['0', 'N кр\n{:,.0f}'.format(chapter_10.N_kr), 'N пл\n{:,.0f}'.format(initial_data.N_pl)], rotation=0)
    plt.yticks([0, S_const_costs[-1] / 1e3, chapter_10.N_kr * chapter_7.P_proizv_plan / 1e3, S_total_costs[-1] / 1e3, chapter_8.Q_plan / 1e3],
               ['0', 'S усл.пост.\n{:,.0f}'.format(S_const_costs[-1] / 1e3), 'Q кр.\n{:,.0f}'.format(chapter_10.Q_kr / 1e3),
                '{:,.0f}\nS сум.'.format(S_total_costs[-1] / 1e3), 'Q пл\n{:,.0f}'.format(chapter_8.Q_plan / 1e3)])
    plt.grid(True)
    # plt.plot([0, N_kr], [N_kr * TS_B_proizv_plan / 1e3, N_kr * TS_B_proizv_plan / 1e3], color='#444444', ls=':')
    plt.plot(N_pl_values, np.array(S_const_costs) / 1e3, label='S усл-пост.', ls=':')
    plt.plot(N_pl_values, np.array(S_variable_costs) / 1e3, label='S перем.', ls=':')
    plt.plot(N_pl_values, np.array(S_total_costs) / 1e3, label='S тек.сум.')
    plt.plot([0, initial_data.N_pl], [0, chapter_8.Q_plan / 1e3], label='Q пл.')
    plt.legend()
    plt.tight_layout()

    memfile = BytesIO()
    plt.savefig(memfile)

    picP = document.add_paragraph()
    picP.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    picP.add_run().add_picture(memfile, width=Cm(17))
    picP.add_run().add_break(WD_BREAK.PAGE)




def main():
    init_styles()
    gen_introduction()
    gen_initial_data()
    gen_1_1()
    gen_1_2()
    gen_1_3()
    gen_1_4()
    gen_1_5()
    gen_1_6()
    gen_1_7()
    gen_1_8()
    gen_1_9()
    gen_1_10()


if __name__ == '__main__':
    import time

    t = time.time()
    main()

    sections = document.sections
    for section in sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(3)
        section.right_margin = Cm(1.5)
    document.save('out.docx')
    print('time: {:,.2f}'.format(time.time() - t))
