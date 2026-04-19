# -*- coding: utf-8 -*-
import io
from odoo import models, _
from odoo.exceptions import UserError
from odoo.tools.misc import get_lang

try:
    import xlsxwriter
except ImportError:
    xlsxwriter = None


class GeneralLedgerXlsx(models.TransientModel):
    """Export XLSX pour le wizard General Ledger."""

    _inherit = 'account.report.general.ledger'

    def action_export_xlsx(self):
        self.ensure_one()
        if xlsxwriter is None:
            raise UserError(_("La bibliothèque Python 'xlsxwriter' est requise."))

        if self.initial_balance and not self.date_from:
            raise UserError(_("You must define a Start Date"))

        data = {'ids': self.env.context.get('active_ids', []),
                'model': self.env.context.get('active_model', 'ir.ui.menu'),
                'form': self.read(['date_from', 'date_to', 'journal_ids',
                                    'target_move', 'company_id',
                                    'display_account', 'initial_balance',
                                    'sortby'])[0]}
        used_context = self._build_contexts(data)
        data['form']['used_context'] = dict(
            used_context, lang=get_lang(self.env).code)

        model = data['model']
        docs = self.env[model].browse(data.get('ids', []))
        accounts = docs if model == 'account.account' else self.env[
            'account.account'].search([])
        report_env = self.env['report.base_accounting_kit.report_general_ledger']
        accounts_res = report_env.with_context(
            data['form'].get('used_context'))._get_account_move_entry(
            accounts, data['form']['initial_balance'],
            data['form']['sortby'], data['form']['display_account'])

        accounts_res.sort(
            key=lambda a: float(a.get('balance') or 0.0), reverse=True)

        journal_codes = [j.code for j in self.env['account.journal'].browse(
            data['form'].get('journal_ids') or [])]
        return self._build_xlsx(data['form'], accounts_res, journal_codes)

    def _build_xlsx(self, data_form, accounts_res, journal_codes):
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        helper = self.env['hilatech.xlsx.report.helper']
        formats = helper._xlsx_get_formats(workbook)

        report_name = _('General Ledger')
        sheet = workbook.add_worksheet(report_name[:31])
        sheet.hide_gridlines(2)

        widths = [12, 10, 16, 28, 22, 14, 14, 14, 18]
        for col, w in enumerate(widths):
            sheet.set_column(col, col, w)
        ncols = len(widths)

        info = helper._xlsx_standard_info_pairs(data_form, journal_codes)
        row = helper._xlsx_write_header(workbook, sheet, formats,
                                         report_name, info, ncols)

        headers = [_('Date'), _('JRNL'), _('Partner'), _('Ref'),
                   _('Move/Entry'), _('Debit'), _('Credit'),
                   _('Balance'), _('Currency')]
        row = helper._xlsx_write_table_header(sheet, row, headers, formats)

        for acc in accounts_res:
            sheet.merge_range(row, 0, row, 4,
                              f"{acc.get('code') or ''} - {acc.get('name') or ''}",
                              formats['level_1'])
            sheet.write_number(row, 5, float(acc.get('debit') or 0.0),
                               formats['level_1_num'])
            sheet.write_number(row, 6, float(acc.get('credit') or 0.0),
                               formats['level_1_num'])
            sheet.write_number(row, 7, float(acc.get('balance') or 0.0),
                               formats['level_1_num'])
            sheet.write(row, 8, '', formats['level_1'])
            row += 1

            for line in acc.get('move_lines') or []:
                sheet.write(row, 0, str(line.get('ldate') or ''),
                            formats['detail'])
                sheet.write(row, 1, line.get('lcode') or '',
                            formats['detail'])
                sheet.write(row, 2, line.get('partner_name') or '',
                            formats['detail'])
                sheet.write(row, 3, line.get('lref') or '',
                            formats['detail'])
                sheet.write(row, 4,
                            f"{line.get('move_name') or ''} / {line.get('lname') or ''}".strip(' /'),
                            formats['detail'])
                sheet.write_number(row, 5, float(line.get('debit') or 0.0),
                                   formats['detail_num'])
                sheet.write_number(row, 6, float(line.get('credit') or 0.0),
                                   formats['detail_num'])
                sheet.write_number(row, 7, float(line.get('balance') or 0.0),
                                   formats['detail_num'])
                sheet.write(row, 8, line.get('currency_code') or '',
                            formats['detail'])
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
