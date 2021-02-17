# -*- coding: utf-8 -*-

import os
from io import BytesIO
from datetime import datetime
from django.conf import settings
from django.http import HttpResponse
from django.contrib import messages
from django.utils.translation import gettext_lazy as _

try:
    from django.template.loaders.app_directories import app_template_dirs
except:
    #from django.template.loaders.app_directories import get_app_template_dirs
    from django.template.utils import get_app_template_dirs
    app_template_dirs = get_app_template_dirs('templates')

def _template_dirs():
    _dirs = list(app_template_dirs)
    try:
        _dirs += settings.TEMPLATES[0]['DIRS']
    except:
        pass
    _dir = os.path.abspath(os.path.dirname(__file__))
    _dirs.append(_dir)
    return _dirs

template_dirs = _template_dirs()

def get_tpl_file(fname):
    for _dir in template_dirs:
        pth = os.path.join(_dir, fname)
        if os.path.exists(pth):
            return pth

class ExportActionMixin(object):
    queryset_name = 'queryset'
    obj_name = 'obj'
    debug = False

    def get_default_payload(self, queryset, list_display, tpl_name=None):
        payload = {self.queryset_name: queryset, 'headers': list_display, 'tpl_name': tpl_name}
        return payload

    def get_extra_payloads(self, queryset, list_display, tpl_name=None):
        payloads = []
        for obj in queryset:
            payload = {self.obj_name: obj, 'tpl_name': tpl_name}
            payloads.append(payload)
        return payloads

    def get_payloads(self, queryset, list_display):
        payloads = []
        payload = self.get_default_payload(queryset, list_display)
        payloads.append(payload)
        return payloads

    def get_export_data(self, queryset, list_display):
        payloads = self.get_payloads(queryset, list_display)
        return self.render(payloads)

    def export(self, request, queryset, list_dispaly, model):
        fname = get_tpl_file(self.tpl)
        if fname is None:
            messages.warning(request, _("Template file '%s' not found." % self.tpl))
        else:
            writer = self.get_writer(fname)
            export_data = self.get_export_data(queryset, list_dispaly)
            response = HttpResponse(export_data, content_type=self.content_type)
            response['Content-Disposition'] = self.get_content_disposition(request, queryset, model)
            return response

    def get_content_disposition(self, request, queryset, model):
        time_str = datetime.now().strftime('%Y-%m-%d-%H-%M')
        filename = "%s-%s.%s" % (model.__name__, time_str, self.file_format)
        content_disposition = 'attachment; filename="%s"' % filename
        return content_disposition

class Xls(ExportActionMixin):
    content_type = 'application/vnd.ms-excel'
    file_format = 'xls'
    desc = _("to xls")

    def get_writer(self, fname):
        from .writer import TplBookWriter
        self.writer =  TplBookWriter(fname, self.debug)
        return self.writer

    def render(self, payloads):
        self.writer.render_book(payloads)
        stream = BytesIO()
        self.writer.save(stream)
        return stream.getvalue()

class Xlsx(ExportActionMixin):
    content_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    file_format = 'xlsx'
    desc = _("to xlsx")

    def get_writer(self, fname):
        from xltpl.writerx import BookWriter
        self.writer = BookWriter(fname, self.debug)
        return self.writer

Xlsx.render = Xls.render

class Docx(ExportActionMixin):
    content_type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    file_format = 'docx'
    desc = _("to docx")

    def get_writer(self, fname):
        from pydocxtpl import DocxWriter
        self.writer = DocxWriter(fname, self.debug)
        return self.writer

    def get_payloads(self, queryset, list_display):
        return self.get_default_payload(queryset, list_display)

    def render(self, payload):
        self.writer.render(payload)
        stream = BytesIO()
        self.writer.save(stream)
        return stream.getvalue()

class XlsDefault(Xls):
    tpl = 'default.xls'

    def get_writer(self, fname):
        from .writer import DefaultBookWriter
        self.writer = DefaultBookWriter(fname)
        return self.writer

    def render(self, payloads):
        self.writer.write_payloads(payloads)
        stream = BytesIO()
        self.writer.save(stream)
        return stream.getvalue()

class XlsxDefault(Xlsx):
    tpl = 'default.xlsx'

    def get_writer(self, fname):
        from .writer import DefaultBookWriterx
        self.writer = DefaultBookWriterx(fname)
        return self.writer

XlsxDefault.render = XlsDefault.render

