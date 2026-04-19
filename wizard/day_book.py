# -*- coding: utf-8 -*-
import io
from datetime import date as date_cls, datetime, timedelta
from odoo import models, _
from odoo.exceptions import UserError

try:
    import xlsxwriter
except ImportError:
    xlsxwriter = None


class DayBookXlsx(models.TransientModel):
    """Export XLSX pour le wizard Day Book."""

    _inherit = 'account.day.book.report'

    def action_export_xlsx(self):
        self.ensure_one()
        if xlsxwriter is None:
            raise UserError(_("La bibliothèque Python 'xlsxwriter' est requise."))

        data = {'ids': self.env.context.get('active_ids', []),
                'model': self.env.context.get('active_model', 'ir.ui.menu'),
                'form': self.read([
                    'date_from', 'date_to', 'journal_ids', 'target_move',
                    'account_ids', 'company_id',
                ])[0]}
        used_context = self._build_contexts(data)
        data['form']['used_context'] = dict(
            used_context, lang=self.env.context.get('lang') or 'en_US')

        active_acc = data['form']['account_ids']
        accounts = (self.env['account.account'].browse(active_acc)
                    if active_acc
                    else self.env['account.account'].search([]))

        form_copy = dict(data['form'])
        form_copy['date_from'] = str(data['form']['date_from'])
        form_copy['date_to'] = str(data['form']['date_to'])

        report_env = self.env[
            'report.base_accounting_kit.day_book_report_template']

        date_start = datetime.strptime(form_copy['date_from'],
                                       '%Y-%m-%d').date()
        date_end = datetime.strptime(form_copy['date_to'],
                                     '%Y-%m-%d').date()
        days = (date_end - date_start).days
        records = []
        for i in range(days + 1):
            pass_date = str(date_start + timedelta(days=i))
            res = report_env.with_context(
                data['form'].get('used_context', {}))._get_account_move_entry(
                accounts, form_copy, pass_date)
            if res['lines']:
                records.append({
                    'date': date_start + timedelta(days=i),
                    'debit': res['debit'],
                    'credit': res['credit'],
                    'balance': res['balance'],
                    'child_lines': res['lines'],
                })

        journal_codes = [j.code for j in self.env['account.journal'].browse(
            data['form'].get('journal_ids') or [])]
        return self._build_day_book_xlsx(data['form'], records, journal_codes)

    def _build_day_book_xlsx(self, data_form, records, journal_codes):
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        helper = self.env['hilatech.xlsx.report.helper']
        formats = helper._xlsx_get_formats(workbook)

        report_name = _('Day Book')
        sheet = workbook.add_worksheet(report_name[:31])
        sheet.hide_gridlines(2)

        widths = [10, 22, 10, 22, 18, 28, 14, 14, 14]
        for col, w in enumerate(widths):
            sheet.set_column(col, col, w)
        ncols = len(widths)

        info = helper._xlsx_standard_info_pairs(data_form, journal_codes)
        row = helper._xlsx_write_header(workbook, sheet, formats,
                                         report_name, info, ncols)

        headers = [_('Date'), _('Account'), _('JRNL'), _('Partner'),
                   _('Ref'), _('Move/Entry'), _('Debit'), _('Credit'),
                   _('Balance')]
        row = helper._xlsx_write_table_header(sheet, row, headers, formats)

        for rec in records:
            header_date = (rec['date'].isoformat()
                           if isinstance(rec['date'], date_cls)
                           else str(rec['date']))
            sheet.merge_range(row, 0, row, ncols - 4,
                              _('Date: %s') % header_date,
                              formats['level_1'])
            sheet.write_number(row, ncols - 3,
                               float(rec.get('debit') or 0.0),
                               formats['level_1_num'])
            sheet.write_number(row, ncols - 2,
                               float(rec.get('credit') or 0.0),
                               formats['level_1_num'])
            sheet.write_number(row, ncols - 1,
                               float(rec.get('balance') or 0.0),
                               formats['level_1_num'])
            row += 1

            lang = self.env.context.get('lang') or 'en_US'

            def _as_text(value):
                if isinstance(value, dict):
                    return value.get(lang) or next(iter(value.values()), '')
                return value or ''

            lines = sorted(rec.get('child_lines') or [],
                           key=lambda l: float(l.get('balance') or 0.0),
                           reverse=True)
            for line in lines:
                sheet.write(row, 0, str(line.get('ldate') or ''),
                            formats['detail'])
                sheet.write(row, 1, _as_text(line.get('accname')),
                            formats['detail'])
                sheet.write(row, 2, _as_text(line.get('lcode')),
                            formats['detail'])
                sheet.write(row, 3, _as_text(line.get('partner_name')),
                            formats['detail'])
                sheet.write(row, 4, _as_text(line.get('lref')),
                            formats['detail'])
                sheet.write(row, 5, _as_text(line.get('move_name')
                            or line.get('lname')), formats['detail'])
                sheet.write_number(row, 6,
                                   float(line.get('debit') or 0.0),
                                   formats['detail_num'])
                sheet.write_number(row, 7,
                                   float(line.get('credit') or 0.0),
                                   formats['detail_num'])
                sheet.write_number(row, 8,
                                   float(line.get('balance') or 0.0),
                                   formats['detail_num'])
                row += 1
            row += 1

        sheet.merge_range(row, 0, row, ncols - 1,
                          _('Generated by Odoo — Hilatech Export Accounting Report'),
                          formats['footer'])

        workbook.close()
        xlsx_bytes = output.getvalue()
        output.close()

        filename = helper._xlsx_build_filename(report_name, data_form)
        return helper._xlsx_build_attachment_action(filename, xlsx_bytes)
