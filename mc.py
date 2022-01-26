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
        # self.vpr.add_row('Настройщик оборудования', 0.05, 60_000)
        # self.vpr.add_row('Складовщик', 0.07, 50_000)
        # self.vpr.add_row('Уборщик', 0.05, 30_000)
        # self.vpr.add_row('Контролёр ОТК', 0.07, 80_000)
        self.vpr.add_row('Настройщик оборудования', 0.07, 90_000)
        self.vpr.add_row('Складовщик', 0.07, 50_000)
        self.vpr.add_row('Уборщик', 0.1, 40_000)
        self.vpr.add_row('Контролёр ОТК', 0.1, 110_000)

        self.sl = PerPercentTable(self.R_opr, True, False)
        self.sl.add_row('Сотрудник', 0.6, 80_000)
        # self.sl.add_row('Генеральный директор', 0, 120_000)
        # self.sl.add_row('HR', 0, 70_000)
        # self.sl.add_row('Менеджер по закупу', 0, 70_000)
        # self.sl.add_row('Менеджер по производству', 0, 85_000)
        # self.sl.add_row('Инженер', 2. / self.R_opr, 85_000)
        # self.sl.add_row('Бухгалтер', 2. / self.R_opr, 80_000)
        # self.sl.add_row('Охраниик', 4. / self.R_opr, 35_000)
        # self.sl.add_row('Логист', 3. / self.R_opr, 50_000)
        # self.sl.add_row('Курьер', 0.5 - 10. / self.R_opr, 45_000)

        self.R_vpr = self.vpr.total
        self.R_sl = self.sl.total

        self.opr_salary = 100_000
        self.opr_extra = 5_000

        self.stimulating_salary_percent = 0.0

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

        self.costs = Value('all', display_name='Затраты')
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
        ft = fuel_energy_costs.add_child(Value('fuel tech', 0, round(fuelt * self.fuel_tech_percent, 2), display_name='Технологическое топливо и энергия'))
        fuel_energy_costs.add_child(
            Value('fuel non tech', fuelt - ft.total, display_name='Нетехнологическое топливо и энергия')
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
        'S_b_proizv',
        'S_b_poln',
        'S_kom_percent',
        'S_kom_const_percent',
        'S_kom',
        'S_B_sum'
    ]

    def __init__(self, initial_data: InitialData, chapter_3: Chapter_3, const=None):
        self.S_b_proizv = chapter_3.S_pr_tek_pl / initial_data.N_pl
        self.S_kom_percent = 0.04
        self.S_kom_const_percent = 0.6
        s = round(self.S_b_proizv * self.S_kom_percent, 2) * initial_data.N_pl
        sc = round(s * self.S_kom_const_percent, 2) if const is None or const['S_kom'] is None else const['S_kom'].const
        sv = round(s - sc, 2) if const is None or const['S_kom'] is None else s * (1 - self.S_kom_const_percent)
        self.S_kom = Value('S_kom', sc, sv, 'Коммерческие затраты')
        self.S_B_sum = Value('S_sum', display_name='Суммарные затраты')
        self.S_B_sum.add_child(chapter_3.costs)
        self.S_B_sum.add_child(self.S_kom)
        self.S_b_poln = Value('S_b_poln', self.S_B_sum.const / initial_data.N_pl, self.S_B_sum.variable / initial_data.N_pl)


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
        self.k_nats = self.profit_before_tax / chapter_4.S_B_sum.total
        self.P_b_poln = round(chapter_4.S_b_poln.total * (1 + self.k_nats), 2)
        self.P_b_perem = round(chapter_4.S_b_poln.variable * (1 + (self.profit_before_tax + chapter_4.S_B_sum.const) / chapter_4.S_B_sum.variable), 2)
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




def main():
    init_styles()
    # gen_introduction()
    # gen_initial_data()
    # gen_1_1()
    # gen_1_2()
    # gen_1_3()
    # gen_1_4()
    # gen_1_5()
    # gen_1_6()
    # gen_1_7()
    # gen_1_8()
    # gen_1_9()


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
