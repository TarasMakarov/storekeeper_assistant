from datetime import datetime

import warehouses as w
import preparation as pr


# search sku in two and more places
# поиск товара, размещенного в двух и более ячейках отгрузки
def a_search_sku_in_two_places(file_at):
    row = 4
    rows = file_at.active.max_row - row
    duplicates_temp = dict()
    for i in range(rows + 1):
        cell_value = file_at.active.cell(row, 2).value
        if isinstance(cell_value, str):
            temp_list = cell_value.split(', ')
        else:
            row += 1
            continue
        sku = file_at.active.cell(row, 4).value.split('.')[0]  # sku without batch
        if sku not in duplicates_temp.keys():
            duplicates_temp[sku] = temp_list
            row += 1
        else:
            for t in temp_list:
                if t not in duplicates_temp[sku]:
                    duplicates_temp[sku].append(t)
            row += 1
    duplicates_result = dict()
    for k, v in duplicates_temp.items():
        if len(v) > 1:
            duplicates_result[k] = v
    return duplicates_result


# search two or more sku in one place
# поиск одного товара, размещенного в двух и более местах
def b_search_two_or_more_sku_in_one_place(file_at):
    row = 4
    col = 1
    rows = file_at.active.max_row - row
    cells_duplicate_temp = dict()
    for i in range(rows + 1):
        cell_value = file_at.active.cell(row, col).value
        if isinstance(cell_value, str):
            temp_list = cell_value.split(', ')
            for t in temp_list:
                cells_duplicate_temp[t] = cells_duplicate_temp.get(t, 0) + 1
            row += 1
        else:
            row += 1
            continue

    cells_duplicate = dict()
    for k, v in cells_duplicate_temp.items():
        if v > 1:
            cells_duplicate[k] = v
    return cells_duplicate


def c_search_cells_empty(file_at):
    row = 4
    rows = file_at.active.max_row - row
    cells_busy = []
    cells_empty_406 = []
    cells_empty_437 = []

    # создается список занятых ячеек
    for i in range(rows + 1):
        cell_value = file_at.active.cell(row, 2).value
        if isinstance(cell_value, str):
            temp_list = cell_value.split(', ')
            cells_busy.extend(temp_list)
            row += 1
        else:
            row += 1

    for c in w.cells_bottom:
        if c not in cells_busy:
            if c[0] in 'ABCDEF':
                cells_empty_437.append(c)
            else:
                cells_empty_406.append(c)
    return cells_empty_406, cells_empty_437


def fill_mistakes(file_at, *args):  # storekeeper
    mistakes_not_found = {'Ошибки не обнаружены': ''}
    operation_was_not_performed = {'Данная операция не выполнялась.': ''}
    duplicates_result = operation_was_not_performed
    wrong_place = operation_was_not_performed
    cells_empty = operation_was_not_performed
    rows_in_cells_empty = 1  # если ниже будет информация

    for arg in args:
        if arg == 'a':
            duplicates_result = a_search_sku_in_two_places(file_at)
            if not duplicates_result:
                duplicates_result = mistakes_not_found
        if arg == 'b':
            wrong_place = b_search_two_or_more_sku_in_one_place(file_at)
            if not wrong_place:
                wrong_place = mistakes_not_found
        if arg == 'c':
            cells_empty = c_search_cells_empty(file_at)
            rows_in_cells_empty = max(len(cells_empty[0]), len(cells_empty[1]))  # если ниже будет информация
            if not cells_empty:
                cells_empty = {'Свободные ячейки отсутствуют.': ''}

    duplicates_result_len = len(duplicates_result)
    wrong_place_len = len(wrong_place)
    mistakes_dict = {'Товар, размещенный в двух и более ячейках отгрузки': duplicates_result_len,
                     'Ячейки, в которых больше одного вида товара': wrong_place_len,
                     'Свободные ячейки': 0}
    report_mistakes = pr.create_report_mistakes_template(mistakes_dict)

    row = 2
    col = 1
    for k, v in duplicates_result.items():
        report_mistakes.active.cell(row=row, column=col).value = k
        report_mistakes.active.cell(row=row, column=col + 1).value = ', '.join(v)
        row += 1

    row += 2
    for key, value in wrong_place.items():
        report_mistakes.active.cell(row=row, column=col).value = key
        report_mistakes.active.cell(row=row, column=col + 1).value = value
        row += 1

    row += 2
    report_mistakes.active.cell(row=row, column=col).value = 406
    report_mistakes.active.cell(row=row, column=col + 1).value = 437
    row_406 = row_437 = row + 1
    if isinstance(cells_empty, dict):
        for k in cells_empty.keys():
            report_mistakes.active.cell(row=row_406, column=col).value = k
    else:
        for c_406 in cells_empty[0]:
            report_mistakes.active.cell(row=row_406, column=col).value = c_406
            row_406 += 1
        for c_437 in cells_empty[1]:
            report_mistakes.active.cell(row=row_437, column=col + 1).value = c_437
            row_437 += 1

    report_mistakes.save(f'Ошибки от {datetime.today().strftime("%d.%m.%y %H_%M_%S")}.xlsx')
