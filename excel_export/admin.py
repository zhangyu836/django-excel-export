# -*- coding: utf-8 -*-

from django import forms
from django.contrib import admin, messages
from django.contrib.admin.helpers import ActionForm
from django.utils.safestring import mark_safe
from django.forms.widgets import Select
from django.utils.translation import gettext_lazy as _

from .action import XlsDefault, XlsxDefault

class wSelect(Select):
    _js_tpl = """
    <script type="text/javascript">   
    (function($) {
    $(document).ready(function() {
        var $actionsSelect, $choicesElement;
        $actionsSelect = $('#changelist-form select[name="action"]');
        $choicesElement = $('#changelist-form select[name="%s"]').parent();
        $actionsSelect.change(function() {
            if ($(this).val() === '%s') {
                $choicesElement.show();
            } else {
                $choicesElement.hide();
            }
            });
        $actionsSelect.change();
        });
    })(django.jQuery);
    </script>
    """
    def render(self, name, value, attrs=None, renderer=None):
        action = '_export'
        output = super().render(name, value, attrs, renderer)
        output += mark_safe(self._js_tpl % (name, action))
        return output

class ChoiceField(forms.ChoiceField):
    widget = wSelect

def export_action_form_factory(exportActions):

    class _ExportActionForm(ActionForm):
        export_action = ChoiceField(
            label=_('Export Choices'), choices=exportActions, required=False)
    _ExportActionForm.__name__ = str('ExportActionForm')

    return _ExportActionForm

class ExportAdmin(admin.ModelAdmin):
    actions = ['_export']
    export_actions = []
    export_fields = []

    def __init__(self, *args, **kwargs):
        choices = []
        export_actions = self.get_export_actions()
        if export_actions:
            for i, export_action in enumerate(export_actions):
                choices.append((str(i), export_action.desc))
        self.action_form = export_action_form_factory(choices)
        super().__init__(*args, **kwargs)

    def _export(self, request, queryset):
        _export_action = request.POST.get('export_action')
        if not _export_action:
            messages.warning(request, _('You must select an export choice.'))
        else:
            export_actions = self.get_export_actions()
            export_action = export_actions[int(_export_action)]()
            if self.export_fields:
                list_display = self.export_fields
            else:
                list_display = self.get_list_display(request)
            return export_action.export(request, queryset, list_display, self.model)

    def get_export_actions(self):
        if self.export_actions:
            return self.export_actions
        else:
            return [XlsxDefault, XlsDefault]

    _export.short_description = _("Export selected")
