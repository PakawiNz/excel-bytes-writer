import re
from collections import OrderedDict

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.writer.excel import save_virtual_workbook

class_mapper = dict(
    font=Font,
    fill=PatternFill,
    alignment=Alignment,
    border=Border,
)

styles = OrderedDict((
    ('red', dict(
        font=dict(color='FF9C0006'),
        fill=dict(fill_type='solid', fgColor='FFFFC7CE'),
    )),
    ('yellow', dict(
        font=dict(color='FF9C6500'),
        fill=dict(fill_type='solid', fgColor='FFFFEB9C'),
    )),
    ('green', dict(
        font=dict(color='FF006100'),
        fill=dict(fill_type='solid', fgColor='FFC6EFCE'),
    )),
    ('clear', dict(
        fill=dict(fill_type='solid', fgColor='FFFFFFFF'),
    )),
    ('bold', dict(
        font=dict(bold=True),
    )),
    ('italic', dict(
        font=dict(italic=True),
    )),
    ('underline', dict(
        font=dict(underline="single"),
    )),
    ('center', dict(
        alignment=dict(horizontal='center'),
    )),
    ('right', dict(
        alignment=dict(horizontal='right'),
    )),
    ('left', dict(
        alignment=dict(horizontal='left'),
    )),
    ('celled', dict(
        border=dict(
            left=Side(style='thin', color='FF000000'),
            right=Side(style='thin', color='FF000000'),
            top=Side(style='thin', color='FF000000'),
            bottom=Side(style='thin', color='FF000000'),
        )
    )),
))


def apply_style(cell, style):
    group = {}
    for class_key in styles.keys():
        if re.search(r'\b{}\b'.format(class_key), style):
            class_style = styles[class_key]
            for style_key, style_value in class_style.items():
                group.setdefault(style_key, {}).update(style_value)

    for key, kwargs in group.items():
        setattr(cell, key, class_mapper[key](**kwargs))


class ExcelBytesWriter:
    def __init__(self, file_name):
        self.wb = Workbook()
        self.ws = self.wb.active
        self.file_name = file_name

        self.row_pointer = 1
        self.col_pointer = 1

    def set_col_size(self, col, size):
        self.ws.column_dimensions[get_column_letter(col)].width = size

    def add_row(self):
        self.row_pointer += 1
        self.col_pointer = 1

    def add_col(self, value='', span=1, style=''):
        self.ws.cell(
            row=self.row_pointer,
            column=self.col_pointer,
            value=value
        )

        if span > 1:
            self.ws.merge_cells(
                start_row=self.row_pointer,
                start_column=self.col_pointer,
                end_row=self.row_pointer,
                end_column=self.col_pointer + span - 1
            )

        if style:
            for col in range(self.col_pointer, self.col_pointer + span):
                apply_style(self.ws.cell(row=self.row_pointer, column=col), style)

        self.col_pointer += span

    def render(self):
        return save_virtual_workbook(self.wb), self.file_name
