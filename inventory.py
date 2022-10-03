# all functions for foreman
from datetime import datetime

import warehouses as w
import design_of_report as dor
import preparation as pr


def add_row_in_inventory(file_at, row, column_warehouse):
    values_of_row = [file_at.active.cell(row, 1).value, file_at.active.cell(row, 2).value,
                     file_at.active.cell(row, 4).value, file_at.active.cell(row, 5).value,
                     file_at.active.cell(row, 6).value, file_at.active.cell(row, column_warehouse - 1).value,
                     file_at.active.cell(row, column_warehouse).value]
    return values_of_row


def fill_rows_of_report(report, values_for_rows):
    row_start = 4  # с этой строки начинается заполнение (выше - шапка отчета)
    for value in values_for_rows:
        if value[0]:
            height = (int(len(value[0]) / 41) + 1) * 15
            report.active.row_dimensions[row_start].height = height
            report.active.cell(row_start, 1).alignment = dor.alignment_wrap_text
        for c in range(1, 8):
            report.active.cell(row_start, c).value = value[c - 1]
        row_start += 1
    return report


def create_file_inventory_cells(file_at, warehouse, cell_start, cell_finish):
    report_inventory_cells = pr.create_report_inventory_template(warehouse, 'ячеек')
    cells_in_warehouse = []
    values_of_second_row = [str(file_at.active.cell(2, column).value) for column in range(1, 11)]
    column_warehouse = values_of_second_row.index(warehouse) + 1  # номер столбца с номером склада, который считают
    row_start = 4  # с 4-ой строки начинаются данные
    rows = file_at.active.max_row
    if warehouse == w.warehouse_406:
        cells_in_warehouse = w.cells_all_406
    elif warehouse == w.warehouse_437:
        cells_in_warehouse = w.cells_all_437
    index_start = cells_in_warehouse.index(cell_start)
    index_finish = cells_in_warehouse.index(cell_finish) + 1
    cells_for_inventory = cells_in_warehouse[index_start:index_finish]
    values_of_cells_for_inventory = []  # список списков со значениями из строк с нужными ячейками
    for _ in range(rows - row_start + 1):
        target_cells = file_at.active.cell(row_start, 2).value
        if target_cells is not None and len(target_cells) < 10 and target_cells in cells_for_inventory:
            values_of_row = add_row_in_inventory(file_at, row_start, column_warehouse)
            values_of_cells_for_inventory.append(values_of_row)
            row_start += 1
        elif target_cells is not None and len(target_cells) > 9:
            target_cells.split(', ')
            for t in target_cells:
                if t in cells_for_inventory:
                    values_of_row = add_row_in_inventory(file_at, row_start, column_warehouse)
                    values_of_cells_for_inventory.append(values_of_row)
            row_start += 1
        else:
            row_start += 1

    report_inventory_cells = fill_rows_of_report(report_inventory_cells, values_of_cells_for_inventory)
    report_inventory_cells = pr.prepare_for_print(report_inventory_cells)
    report_inventory_cells.save(
        f'Инвентаризация ячеек {datetime.today().strftime("%d.%m.%y %H_%M_%S")} ({warehouse}).xlsx')


# file_supervisor - файл с ошибками
def create_file_inventory_supervisor(file_at, file_supervisor, warehouse, date_start, date_finish):
    report_inventory_mistakes = pr.create_report_inventory_template(warehouse, 'ошибок')
    values_of_second_row = [str(file_at.active.cell(2, column).value) for column in range(1, 11)]
    column_warehouse = values_of_second_row.index(warehouse) + 1  # номер столбца с номером склада, который считают
    sheet_active = file_supervisor["СБОРКА"]
    file_supervisor.active = sheet_active
    rows = sheet_active.max_row
    sku_mistake = []
    for i in range(rows, 0, -1):
        date_mistake = sheet_active.cell(i, 2).value
        if isinstance(date_mistake, datetime):
            date_mistake = sheet_active.cell(i, 2).value.date()
        else:
            date_mistake = datetime(1, 1, 1).date()
        sku = sheet_active.cell(i, 3).value
        if date_start <= date_mistake <= date_finish:
            if sku:
                sku_mistake.append(str(sku))
    values_of_cells_for_inventory = []  # список списков со значениями из строк с нужными ячейками
    row_start = 4
    sku_at = file_at.active.cell(row_start, 4).value
    while sku_at:
        sku_at = sku_at.split('.')[0]
        if sku_at in sku_mistake and file_at.active.cell(row_start, column_warehouse).value > 0:
            values_of_row = add_row_in_inventory(file_at, row_start, column_warehouse)
            values_of_cells_for_inventory.append(values_of_row)
        row_start += 1
        sku_at = file_at.active.cell(row_start, 4).value

    report_inventory_mistakes = fill_rows_of_report(report_inventory_mistakes, values_of_cells_for_inventory)
    report_inventory_mistakes = pr.prepare_for_print(report_inventory_mistakes)
    report_inventory_mistakes.save(
        f'Инвентаризация ошибок {datetime.today().strftime("%d.%m.%y %H_%M_%S")} ({warehouse}).xlsx')
