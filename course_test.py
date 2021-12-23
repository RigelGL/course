from course import *


def main():
    material_costs_per_item = 2050
    N = 45000

    material_costs_raw_items = Value('Материальные затраты', 0, N * material_costs_per_item)
    m_base = material_costs_raw_items.total

    material_costs_raw_items.add_child(Value('main', 0, m_base))
    material_costs_raw_items.add_child(Value('helper', 0, round(m_base * 0.05, 2)))
    material_costs_raw_items.add_child(Value('move save', 0, round(m_base * 0.12, 2)))
    material_costs_raw_items.add_child(Value('inventory', round(m_base * 0.03, 2), 0))
    fuelt = round(m_base * 0.55, 2)
    fuel_energy_costs = material_costs_raw_items.add_child(Value('fuel total'))
    ft = fuel_energy_costs.add_child(Value('tech', 0, round(fuelt * 0.7, 2)))
    fuel_energy_costs.add_child(Value('non tech', fuelt - ft.total))

    # print(material_costs_raw_items)

    table = Table('count', 'price')
    table.add_row(3, 400)
    table.add_row(2, 800)
    table.add_row(4, 200)

    # print(table.calculate_sum(lambda x: x['price'] * x['count']))

    p = PercentTable()
    p.add_row(material_costs_raw_items)
    # print(p)

    ppt = PerPercentTable(100_000., normalize=True)
    ppt.add_row(22, 'ОПФ')
    ppt.add_row(5.1, 'ФОМС')
    ppt.add_row(2.9, 'ФСС')
    ppt.add_row(4, 'Нснп')
    print(ppt)


if __name__ == '__main__':
    main()