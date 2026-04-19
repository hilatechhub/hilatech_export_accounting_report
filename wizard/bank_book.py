# -*- coding: utf-8 -*-
import io
from odoo import models, _
from odoo.exceptions import UserError

try:
    import xlsxwriter
except ImportError:
    xlsxwriter = None


class BankBookXlsx(models.TransientModel):
    """Export XLSX pour le wizard Bank Book."""

    _inherit = 'account.bank.book.report'

    def action_export_xlsx(self):
        self.ensure_one()
        if xlsxwriter is None:
            raise UserError(_("La bibliothèque Python 'xlsxwriter' est requise."))
        if self.initial_balance and not self.date_from:
            raise UserError(_("You must choose a Start Date"))

        data = {'ids': self.env.context.get('active_ids', []),
                'model': self.env.context.get('active_model', 'ir.ui.menu'),
                'form': self.read([
                    'date_from', 'date_to', 'journal_ids', 'target_move',
                    'display_account', 'account_ids', 'sortby',
                    'initial_balance', 'company_id',
                ])[0]}
        used_context = self._build_contexts(data)
        data['form']['used_context'] = dict(
            used_context, lang=self.env.context.get('lang') or 'en_US')

        report_env = self.env['report.base_accounting_kit.report_bank_book']
        account_ids = data['form']['account_ids']
        accounts = self.env['account.account'].browse(account_ids)
        if not accounts:
            journals = self.env['account.journal'].search(
                [('type', '=', 'bank')])
            accounts = self.env['account.account'].browse(
                [j.default_account_id.id for j in journals
                 if j.default_account_id])

        accounts_res = report_env.with_context(
            data['form'].get('used_context', {}))._get_account_move_entry(
            accounts, data['form'].get('initial_balance', True),
            data['form'].get('sortby', 'sort_date'), 'movement')

        journal_codes = [j.code for j in self.env['account.journal'].browse(
            data['form'].get('journal_ids') or [])]
        return self._build_bank_cash_xlsx(
            _('Bank Book'), data['form'], accounts_res, journal_codes)

    def _build_bank_cash_xlsx(self, report_name, data_form, accounts_res,
                               journal_codes):
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        helper = self.env['hilatech.xlsx.report.helper']
        formats = helper._xlsx_get_formats(workbook)

        sheet = workbook.add_worksheet(report_name[:31])
        sheet.hide_gridlines(2)

        widths = [12, 10, 22, 18, 28, 14, 14, 14]
        for col, w in enumerate(widths):
            sheet.set_column(col, col, w)
        ncols = len(widths)

        info = helper._xlsx_standard_info_pairs(data_form, journal_codes)
        info.insert(1, (_('Sort By'),
                        _('Date') if data_form.get('sortby') == 'sort_date'
                        else _('Journal & Partner')))
        row = helper._xlsx_write_header(workbook, sheet, formats,
                                         report_name, info, ncols)

        headers = [_('Date'), _('JRNL'), _('Partner'), _('Ref'),
                   _('Move/Entry'), _('Debit'), _('Credit'), _('Balance')]
        row = helper._xlsx_write_table_header(sheet, row, headers, formats)

        accounts_res.sort(key=lambda a: float(a.get('balance') or 0.0),
                          reverse=True)

        for account in accounts_res:
            label = f"{account.get('code') or ''} {account.get('name') or ''}".strip()
            sheet.merge_range(row, 0, row, ncols - 4, label,
                              formats['level_1'])
            sheet.write_number(row, ncols - 3,
                               float(account.get('debit') or 0.0),
                               formats['level_1_num'])
            sheet.write_number(row, ncols - 2,
                               float(account.get('credit') or 0.0),
                               formats['level_1_num'])
            sheet.write_number(row, ncols - 1,
                               float(account.get('balance') or 0.0),
                               formats['level_1_num'])
            row += 1

            for line in account.get('move_lines') or []:
                sheet.write(row, 0, str(line.get('ldate') or ''),
                            formats['detail'])
                sheet.write(row, 1, line.get('lcode') or '',
                            formats['detail'])
                sheet.write(row, 2, line.get('partner_name') or '',
                            formats['detail'])
                sheet.write(row, 3, line.get('lref') or '',
                            formats['detail'])
                sheet.write(row, 4, line.get('move_name')
                            or line.get('lname') or '', formats['detail'])
                sheet.write_number(row, 5,
                                   float(line.get('debit') or 0.0),
                                   formats['detail_num'])
                sheet.write_number(row, 6,
                                   float(line.get('credit') or 0.0),
                                   formats['detail_num'])
                sheet.write_number(row, 7,
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
