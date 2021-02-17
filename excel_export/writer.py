# -*- coding: utf-8 -*-

from datetime import date, time, datetime
from decimal import Decimal
from six import text_type
from xltpl.base import BookBase, SheetBase
from xltpl.basex import BookBase as BookBasex, SheetBase as SheetBasex
from xltpl.pos import Pos


class SheetMixin():
    types = ['', text_type, int, float, bool, Decimal, time, date, datetime]

    def write_headers(self, rdrowx, rdcolx, headers):
        self.pos.next_row()
        for attr in headers:
            wtrowx, wtcolx = self.pos.next_cell()
            self.cell(rdrowx, rdcolx, wtrowx, wtcolx, attr.upper())

    def write_queryset(self, rdrowx, queryset, headers):
        colx_list = self.get_colx_list(headers, queryset[0])
        for obj in queryset:
            self.pos.next_row()
            for index,attr in enumerate(headers):
                wtrowx, wtcolx = self.pos.next_cell()
                field_value = getattr(obj, attr)
                rdcolx = colx_list[index]
                if rdcolx == self.index_base :
                    field_value = text_type(field_value)
                self.cell(rdrowx, rdcolx, wtrowx, wtcolx, field_value)

    def write_payload(self, payload):
        headers = payload.get('headers')
        queryset = payload.get('queryset')
        if not headers or not queryset:
            return
        self.pos.next_row()
        self.write_queryset(self.index_base + 1, queryset, headers)
        self.pos.set_mins(self.index_base, self.index_base)
        self.write_headers(self.index_base, self.index_base, headers)

    def get_colx_list(self, headers, obj):
        colx_list = []
        for attr in headers:
            field_value = getattr(obj, attr)
            fty = type(field_value)
            try:
                index = self.types.index(fty)
            except:
                index = 0
            colx_list.append(index + self.index_base)
        return colx_list

class SheetWriter(SheetMixin, SheetBase):

    def __init__(self, bookwriter, rdsheet, sheet_name):
        SheetBase.__init__(self, bookwriter, rdsheet, sheet_name)
        self.index_base = 0
        self.pos = Pos(self.index_base, self.index_base)

class SheetWriterx(SheetMixin, SheetBasex):

    def __init__(self, bookwriter, rdsheet, sheet_name):
        SheetBasex.__init__(self, bookwriter, rdsheet, sheet_name)
        self.index_base = 1
        self.pos = Pos(self.index_base, self.index_base)

class BookMixin():

    def __init__(self, fname):
        self.load(fname)

    def write_payloads(self, payloads):
        for payload in payloads:
            idx = self.get_tpl_idx(payload)
            sheet_name = self.get_sheet_name(payload)
            writer = self.get_sheet_writer(sheet_name, idx)
            writer.write_payload(payload)

class DefaultBookWriter(BookMixin, BookBase):
    sheet_class = SheetWriter

    def create_wtbook(self):
        if not hasattr(self, 'wtbook'):
            self.create_workbook()

    def get_sheet_writer(self, sheet_name, rdsheet_idx=0):
        self.create_wtbook()
        if sheet_name is None:
            sheet_name = 'XLSheet%d' % len(self.wtbook.wtsheet_names)
        rdsheet = self.rdsheet_list[rdsheet_idx]
        writer = self.sheet_class(self, rdsheet, sheet_name)
        return writer

    def write_payloads(self, payload):
        self.create_wtbook()
        BookMixin.write_payloads(self, payload)

    def save(self, fname):
        self.wtbook.save(fname)
        del self.wtbook

class DefaultBookWriterx(BookMixin, BookBasex):
    sheet_class = SheetWriterx

    def get_sheet_writer(self, sheet_name, rdsheet_idx=0):
        if sheet_name is None:
            sheet_name = 'XLSheet%d' % len(self.workbook._sheets)
        rdsheet = self.rdsheet_list[rdsheet_idx]
        writer = self.sheet_class(self, rdsheet, sheet_name)
        return writer

    def save(self, fname):
        self.workbook.save(fname)
        for sheet in self.workbook.worksheets:
            self.workbook.remove(sheet)

from xltpl.writer import BookWriter
class TplBookWriter(BookWriter):

    def save(self, fname):
        if self.wtbook is not None:
            self.wtbook.save(fname)
            del self.wtbook

