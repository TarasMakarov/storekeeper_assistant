"""
В этом файле располагаются объекты для оформления результатов работы программы (результирующих файлов .xlsx)
Стили шрифта:
    - font_row_1 - для наименования(заголовка) отчета
    - font_row_2 - для наименования колонок отчета
"""
from openpyxl.styles import (
    Alignment, Border, Side, Font
)

font_row_1 = Font(
    name='Calibri',
    size=14,
    bold=True,
)

font_row_2 = Font(
    name='Calibri',
    size=11,
    bold=True,
)

alignment = Alignment(
    horizontal='center',
    vertical='center',
)

alignment_wrap_text = Alignment(
   wrap_text=True
)

alignment_center_wrap_text = Alignment(
    horizontal='center',
    vertical='center',
    wrap_text=True
)


side = Side(
    border_style="thin",
    color="000000"
)

border = Border(
    left=side,
    right=side,
    top=side,
    bottom=side
)
