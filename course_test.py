from course import *
from math import ceil


def calc_mat_costs(n, fot, fot_safe, const=None):
    material_costs_per_item = 2050

    costs = Value('all', display_name='Затраты')

    material_costs_raw_items = Value('material', 0, n * material_costs_per_item, display_name='Материальные затраты')
    m_base = material_costs_raw_items.total

    material_costs_raw_items.add_child(Value('main', 0, m_base, display_name='Основные материалы'))
    material_costs_raw_items.add_child(Value('helper', 0, round(m_base * 0.05, 2), display_name='Вспомогательные материалы'))
    material_costs_raw_items.add_child(Value('move save', 0, round(m_base * 0.12, 2), display_name='Транспортно-заготовительные расходы'))
    material_costs_raw_items.add_child(
        Value('inventory', round(m_base * 0.03, 2), display_name='Инструменты, инвентарь') if const is None or const['inventory'] is None else const['inventory'])
    fuelt = round(m_base * 0.55, 2)
    fuel_energy_costs = material_costs_raw_items.add_child(Value('fuel total', display_name='Топливо и энергия'))
    ft = fuel_energy_costs.add_child(Value('tech', 0, round(fuelt * 0.7, 2), display_name='Технологическое'))
    fuel_energy_costs.add_child(Value('non tech', fuelt - ft.total, display_name='Нетехнологическое') if const is None or const['non tech'] is None else const['non tech'])
    costs.add_child(material_costs_raw_items)

    costs.add_child(fot)
    costs.add_child(fot_safe)

    costs.add_child(Value('amortisation', const=1_150_000))

    costs.add_child(Value('extra', costs.const * 0.12 + costs['amortisation'].variable * 0.07, costs.variable * 0.12 + costs['amortisation'].const * 0.07))

    return costs


def calc_fot(n, fix_vpr=None):
    opr = ceil(n / 1000)

    ppt = PerPercentTable(int(opr * 0.5) if fix_vpr is None else fix_vpr, minimum_is_one=True, normalize=True)
    ppt.add_row('vpr 1', 30, 30_000)
    ppt.add_row('vpr 2', 20, 40_000)
    ppt.add_row('vpr 3', 50, 60_000)

    print(ppt)
    c = ppt.calc_sum(lambda x, y: x * y)
    print(c)
    return Value('fot', c, opr * 100_000)


def main():
    m_fot = calc_fot(45_000)
    m_fot_safe = Value('fot_safe', 0, 0)
    m = calc_mat_costs(45_000, m_fot, m_fot_safe)
    print(m)

    c = CalculateTable([450, 1800, 2700, 18000, 35000, 45000], lambda x: calc_mat_costs(x, m_fot, m_fot_safe, m).head())

    print(c)


if __name__ == '__main__':
    main()
