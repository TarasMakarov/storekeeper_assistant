from datetime import datetime

import preparation as pr


def create_report(file_bc, file_at, maximum, warehouse):
    report_refill_warehouse = pr.create_report_refill_warehouse_template(warehouse)
    column_end = file_bc.active.max_column
    values_of_column = [str(file_bc.active.cell(4, column).value) for column in range(1, column_end + 1)]
    warehouse_storage_column = values_of_column.index(warehouse)  # номер столбца с номером склада хранения
    warehouse_shipping_column = values_of_column.index(warehouse) + 1  # номер столбца с номером склада отгрузки
    row_start_file_bc = 7  # с 7-ой строки начинаются данные в файле "Текущие остатки"
    row_start_report = 3  # внесение данных в отчет с 3-ей строки
    while file_bc.active.cell(row_start_file_bc, 2).value:
        warehouse_shipping_amount = file_bc.active.cell(row_start_file_bc, warehouse_shipping_column).value
        if warehouse_shipping_amount is None or warehouse_shipping_amount <= maximum:
            warehouse_storage_amount = file_bc.active.cell(row_start_file_bc, warehouse_storage_column).value
            if warehouse_storage_amount:
                report_refill_warehouse.active.cell(row_start_report, 1).value = \
                    file_bc.active.cell(row_start_file_bc, 2).value
                report_refill_warehouse.active.cell(row_start_report, 2).value = \
                    file_bc.active.cell(row_start_file_bc, 3).value
                report_refill_warehouse.active.cell(row_start_report, 3).value = \
                    file_bc.active.cell(row_start_file_bc, warehouse_storage_column).value
                report_refill_warehouse.active.cell(row_start_report, 4).value = \
                    file_bc.active.cell(row_start_file_bc, warehouse_shipping_column).value
                row_start_report += 1
        row_start_file_bc += 1

    row_start_report = 3
    rows_in_report_refill_warehouse = report_refill_warehouse.active.max_row

    sku_for_replenish = {report_refill_warehouse.active.cell(row, 1).value: row for row in
                         range(row_start_report, rows_in_report_refill_warehouse + 1)}

    row_start_file_at = 4
    sku_current = file_at.active.cell(row_start_file_at, 4).value
    warehouse_storage_place = 'пусто'
    warehouse_shipping_place = ['пусто']
    while sku_current:
        sku_current = sku_current.split('.')[0]
        if sku_current not in sku_for_replenish.keys():
            row_start_file_at += 1
            sku_current = file_at.active.cell(row_start_file_at, 4).value
        else:
            warehouse_storage_cell = file_at.active.cell(row_start_file_at, 1).value
            if warehouse_storage_cell and warehouse_storage_place == 'пусто':
                warehouse_storage_place = warehouse_storage_cell.split(', ')[0]
            warehouse_shipping_cell = file_at.active.cell(row_start_file_at, 2).value
            if warehouse_shipping_cell:
                warehouse_shipping_cell_values = warehouse_shipping_cell.split(', ')
                for w in warehouse_shipping_cell_values:
                    if w not in warehouse_shipping_place:
                        warehouse_shipping_place.append(w)
            sku_next = file_at.active.cell(row_start_file_at + 1, 4).value
            if sku_next:
                sku_next = sku_next.split('.')[0]
            if sku_next != sku_current:
                report_refill_warehouse.active.cell(sku_for_replenish[sku_current], 5).value = warehouse_storage_place
                if len(warehouse_shipping_place) > 1:
                    report_refill_warehouse.active.cell(sku_for_replenish[sku_current], 6).value = \
                        ', '.join(warehouse_shipping_place[1:])
                else:
                    report_refill_warehouse.active.cell(sku_for_replenish[sku_current], 6).value = \
                        warehouse_shipping_place[0]
                warehouse_storage_place = 'пусто'
                warehouse_shipping_place = ['пусто']
            sku_current = sku_next
            row_start_file_at += 1

    pr.prepare_for_print(report_refill_warehouse)

    values_of_cells_for_report_refill = []
    row_start_report = 3
    value = report_refill_warehouse.active.cell(row_start_report, 1).value
    while value:
        values_of_row = [report_refill_warehouse.active.cell(row_start_report, 1).value,
                         report_refill_warehouse.active.cell(row_start_report, 2).value,
                         report_refill_warehouse.active.cell(row_start_report, 3).value,
                         report_refill_warehouse.active.cell(row_start_report, 4).value,
                         report_refill_warehouse.active.cell(row_start_report, 5).value,
                         report_refill_warehouse.active.cell(row_start_report, 6).value]
        values_of_cells_for_report_refill.append(values_of_row)
        row_start_report += 1
        value = report_refill_warehouse.active.cell(row_start_report, 1).value

    values_of_cells_for_report_refill.sort(key=lambda c: c[5])

    row_start_report = 3
    for i in range(len(values_of_cells_for_report_refill)):
        for j in range(1, 7):
            report_refill_warehouse.active.cell(row_start_report, j).value = values_of_cells_for_report_refill[i][j - 1]
        row_start_report += 1

    report_refill_warehouse.save(f'Отчет от {datetime.today().strftime("%d.%m.%y %H_%M_%S")} ({warehouse}).xlsx')
