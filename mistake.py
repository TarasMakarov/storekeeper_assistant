import warehouses as w


# search sku in two and more places
# поиск товара, размещенного в двух и более ячейках отгрузки
def search_sku_in_two_places(file_at):
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
def search_two_or_more_sku_in_one_place(file_at):
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


def search_cells_empty(file_at):
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
