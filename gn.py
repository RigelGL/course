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

    dp(f'4. Инструменты, инвентарь  (принято за {fn(chapter_3.inventory_percent * 100)}% от *)')
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
    dp('Оборотные  средства участвуют в одном производственном цикле и полностью переносят свою стоимость на готовую продукцию, и '
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
            ['М_{мат\\ i},\\ К_{комп\\ j}', 'норма расхода i-го материала и j-го вида комплектующих изделий на одно  изготавливаемое изделие в стоимостном выражении, руб./шт.'],
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

    dp('Затраты  на материалы и комплектующие изделия:')
    add_formula('S_{мат\\ и\\ комп} =' + f'{fn(chapter_3.S_mat_i_comp)}\\ [руб.]')

    dp('Коэффициент нарастания затрат  (условно принимается равномерное нарастание затрат)')
    add_formula(
        'k_{нз} = \\frac{S_{мат\\ и\\ комп} + S_{Б_{произв}}}{2 S_{Б_{произв}}} = \\frac{' +
        f'{fn(chapter_3.S_mat_i_comp)} + {fn(chapter_4.S_b_proizv)} }}{{ 2 \\cdot {fn(chapter_4.S_b_proizv)} }} = {fn(chapter_5.k_nz)}')

    document.add_page_break()

    dp('Производственный цикл (кален. дни) – отрезок времени между началом и окончанием производственного процесса '
       'изготовления одного изделия, включающий время технологических операций; время подготовительно-заключительных операций; '
       'длительность естественных процессов и вспомогательных операций; время межоперационных и междусменных перерывов; '
       'время ожидания обработки при передаче изделий на рабочие места по партиям')

    add_formula_with_description('T_ц = \\frac{\\sum^{m}_{i=1}{t_{техн\\ i}}\\gamma_ц}{C D}\\frac{T_{пл}}{Т_{пл} - B}', [
        ['\\gamma_ц', f'cоотношение между производственным циклом и суммарной технологической трудоемкостью изготовления изделия, принято за {fn(chapter_5.gamma_cycle, 0)}']
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
    p = add_formula(
        'K_{об._{гот.прод.}} = \\frac{S_{Б_{произв}} N_{пл}}{Т_{пл}}t_{реал} = \\frac{' +
        f'{fn(chapter_4.S_b_proizv)} \\cdot {fn(initial_data.N_pl)} }}{{ {chapter_1.T_pl} }} \\cdot {chapter_5.t_real} = {fn(chapter_5.K_ob_got_prod)}\\ [руб.]')
    p.add_run().add_break(WD_BREAK.PAGE)

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
    dp('Условно в работе вступительный  баланс (составляемый на момент возникновения предприятия) совпадает с текущим балансом, составляемым на начало отчетного периода.')
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
    dp(f'Долгосрочные заёмные средства: {fn(chapter_6.doldosroch_zaemn_sredstva_percent * 100)}')
    dp(f'Краткосрочные заёмные средства: {fn(chapter_6.kratkosroch_zaemn_sredstva_percent * 100)}')
    g = 1 - chapter_6.kratkosroch_zaemn_sredstva_percent - chapter_6.doldosroch_zaemn_sredstva_percent
    dp(f'Прочие краткосрочные обязательства: {fn(g * 100)}').add_run().add_break(WD_BREAK.PAGE)

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
    dp('Прибыль до налогообложения определим через удельный вес чистой прибыли в общей сумме прибыли до налогообложения 0,8  (уточнять ежегодно). '
       'Для этого планируемую (желаемую) чистую прибыль зададим в пределах 20-60% стоимости собственного капитала, руб./год')

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
        dp('Возможно частичное погашение краткосрочных заёмных средств для плановой суммы денежных средств (альнейшая генерация неверна):').runs[0].font.color.rgb = RGBColor(255, 0, 0)
        add_formula('K_{ден.ср.конец\\ план} = K_{ден.ср.\\ план} - S_{кр.заёмн.ср.} = ' + f'{fn(500_000)} \\ [руб.]')
        add_formula('S_{кр.заёмн.ср.\\ план} = ' + f'{fn(chapter_9.S_kratkosroch_zaem_sredstva_konez_plan)} \\ [руб.]')
    else:
        dp('Погашение краткосрочных заёмных средств невозможно для плановой суммы денежных средств, дальнейшая генерация неверна.').runs[0].font.color.rgb = RGBColor(255, 0, 0)

    if chapter_9.valid_to_cope_kz_fact == 'full':
        dp('Возможно полное погашение краткосрочных заёмных средств для фактической суммы денежных средств (дальнейшая генерация неверна):').runs[0].font.color.rgb = RGBColor(255, 0, 0)
        add_formula('K_{ден.ср.конец\\ факт} = K_{ден.ср.\\ факт} - S_{кр.заёмн.ср.} = ' +
                    f'{fn(chapter_9.K_den_sr_fact)} - {fn(chapter_6.active_passive.kratkosroch_zaem_sredstva)} = {fn(chapter_9.K_den_sr_konez_fact)}\\ [руб.]',
                    style=formula_style_12)
    elif chapter_9.valid_to_cope_kz_fact == 'part':
        dp('Возможно частичное погашение краткосрочных заёмных средств для фактической суммы денежных средств:')
        add_formula('K_{ден.ср.конец\\ факт} = K_{ден.ср.\\ факт} - S_{кр.заёмн.ср.} = ' + f'{fn(500_000)} \\ [руб.]')
        add_formula('S_{кр.заёмн.ср.\\ факт} = ' + f'{fn(chapter_9.S_kratkosroch_zaem_sredstva_konez_fact)} \\ [руб.]')
    else:
        dp('Погашение краткосрочных заёмных средств невозможно для фактической суммы денежных средств, дальнейшая генерация неверна.').runs[0].font.color.rgb = RGBColor(255, 0, 0)

    document.add_page_break()

    dp('Таблица 9.1, плановый бухгалтерский баланс на конец периода', table_name_text)
    add_active_passive_table(chapter_9.active_passive_plan)

    document.add_page_break()

    dp('Таблица 9.2, фактический бухгалтерский баланс на конец периода', table_name_text)
    add_active_passive_table(chapter_9.active_passive_fact)
    document.add_page_break()
