# -*- coding: utf-8 -*-
import io
from odoo import fields, models, _
from odoo.exceptions import UserError

try:
    import xlsxwriter
except ImportError:
    xlsxwriter = None


class FinancialReportXlsx(models.TransientModel):
    """Ajoute l'export XLSX au wizard Profit and Loss / Balance Sheet.

    Réutilise intégralement la logique `get_account_lines` du wizard
    `financial.report` de base_accounting_kit pour garantir que les chiffres
    exportés sont identiques à ceux du PDF.
    """

    _inherit = 'financial.report'

    @staticmethod
    def _line_key(line):
        return line['a_id'] if line['type'] == 'account' else line['id']

    def _sort_lines_by_balance_desc(self, report_lines):
        """Trie les lignes par solde décroissant en préservant la
        hiérarchie: chaque parent reste au-dessus de ses enfants, et les
        frères/sœurs sous un même parent sont ordonnés du plus grand solde
        au plus petit."""
        lines_by_key = {self._line_key(l): l for l in report_lines}
        children = {}
        roots = []
        for line in report_lines:
            parent_id = line.get('parent')
            if parent_id and parent_id in lines_by_key:
                children.setdefault(parent_id, []).append(line)
            else:
                roots.append(line)

        def balance_of(line):
            return float(line.get('balance') or 0.0)

        for key in children:
            children[key].sort(key=balance_of, reverse=True)
        roots.sort(key=balance_of, reverse=True)

        ordered = []

        def walk(line):
            ordered.append(line)
            for child in children.get(self._line_key(line), []):
                walk(child)

        for root in roots:
            walk(root)
        return ordered

    def action_export_xlsx(self):
        self.ensure_one()
        if xlsxwriter is None:
            raise UserError(_(
                "La bibliothèque Python 'xlsxwriter' est requise pour "
                "exporter en Excel. Installez-la avec: pip install xlsxwriter"
            ))

        data = {'ids': self.env.context.get('active_ids', []),
                'model': self.env.context.get('active_model', 'ir.ui.menu'),
                'form': self.read([
                    'date_from', 'enable_filter', 'debit_credit', 'date_to',
                    'account_report_id', 'target_move', 'view_format',
                    'company_id',
                ])[0]}
        used_context = self._build_contexts(data)
        data['form']['used_context'] = dict(
            used_context, lang=self.env.context.get('lang') or 'en_US')

        report_lines = self.get_account_lines(data['form'])

        def set_report_level(rec):
            level = 1
            if not rec['parent']:
                return level
            for line in report_lines:
                key = 'a_id' if line['type'] == 'account' else 'id'
                if line[key] == rec['parent']:
                    return level + set_report_level(line)
            return level

        for item in report_lines:
            item['balance'] = round(item['balance'], 2)
            if not item['parent']:
                item['level'] = 1
            else:
                item['level'] = set_report_level(item)

        report_lines = self._sort_lines_by_balance_desc(report_lines)

        report_name = data['form']['account_report_id'][1]
        debit_credit = data['form']['debit_credit']

        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        helper = self.env['hilatech.xlsx.report.helper']
        formats = helper._xlsx_get_formats(workbook)

        sheet = workbook.add_worksheet(report_name[:31])
        sheet.hide_gridlines(2)

        ncols = 4 if debit_credit else 2
        sheet.set_column(0, 0, 55)
        if debit_credit:
            sheet.set_column(1, 3, 18)
        else:
            sheet.set_column(1, 1, 20)

        target_label = (_('All Posted Entries')
                        if data['form']['target_move'] == 'posted'
                        else _('All Entries'))
        info_pairs = [
            (_('Report'), report_name),
            (_('Target Moves'), target_label),
            (_('Start Date'),
             str(data['form']['date_from']) if data['form']['date_from'] else '-'),
            (_('End Date'),
             str(data['form']['date_to']) if data['form']['date_to'] else '-'),
            (_('Printed on'),
             fields.Date.context_today(self).isoformat()),
        ]
        row = helper._xlsx_write_header(workbook, sheet, formats,
                                         report_name, info_pairs, ncols)

        sheet.write(row, 0, _('Name'), formats['table_header'])
        if debit_credit:
            sheet.write(row, 1, _('Debit'), formats['table_header'])
            sheet.write(row, 2, _('Credit'), formats['table_header'])
            sheet.write(row, 3, _('Balance'), formats['table_header'])
        else:
            sheet.write(row, 1, _('Balance'), formats['table_header'])
        row += 1

        for line in report_lines:
            level = line.get('level', 0)
            if level == 0:
                continue
            name = line.get('name', '')
            indent = '    ' * max(level - 1, 0)
            label = f'{indent}{name}'

            if level == 1:
                fmt_text, fmt_num = formats['level_1'], formats['level_1_num']
            elif level == 2:
                fmt_text, fmt_num = formats['level_2'], formats['level_2_num']
            elif level == 3:
                fmt_text, fmt_num = formats['level_3'], formats['level_3_num']
            else:
                fmt_text, fmt_num = formats['detail'], formats['detail_num']

            sheet.write(row, 0, label, fmt_text)
            if debit_credit:
                sheet.write_number(row, 1,
                                   float(line.get('debit') or 0.0), fmt_num)
                sheet.write_number(row, 2,
                                   float(line.get('credit') or 0.0), fmt_num)
                sheet.write_number(row, 3,
                                   float(line.get('balance') or 0.0), fmt_num)
            else:
                sheet.write_number(row, 1,
                                   float(line.get('balance') or 0.0), fmt_num)
            row += 1

        row += 1
        sheet.merge_range(row, 0, row, ncols - 1,
                          _('Generated by Odoo — Hilatech Export Accounting Report'),
                          formats['footer'])

        workbook.close()
        xlsx_bytes = output.getvalue()
        output.close()

        safe_name = report_name.replace(' ', '_').replace('/', '_')
        date_suffix = ''
        if data['form']['date_from'] or data['form']['date_to']:
            date_suffix = (f"_{data['form']['date_from'] or ''}"
                           f"_{data['form']['date_to'] or ''}")
        filename = f'{safe_name}{date_suffix}.xlsx'
        return helper._xlsx_build_attachment_action(filename, xlsx_bytes)
