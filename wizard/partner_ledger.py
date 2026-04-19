# -*- coding: utf-8 -*-
import io
from odoo import models, _
from odoo.exceptions import UserError
from odoo.tools.misc import get_lang

try:
    import xlsxwriter
except ImportError:
    xlsxwriter = None


class PartnerLedgerXlsx(models.TransientModel):
    """Export XLSX pour le wizard Partner Ledger."""

    _inherit = 'account.report.partner.ledger'

    def action_export_xlsx(self):
        self.ensure_one()
        if xlsxwriter is None:
            raise UserError(_("La bibliothèque Python 'xlsxwriter' est requise."))

        data = {'ids': self.env.context.get('active_ids', []),
                'model': self.env.context.get('active_model', 'ir.ui.menu'),
                'form': self.read(['date_from', 'date_to', 'journal_ids',
                                    'target_move', 'company_id',
                                    'result_selection', 'reconciled',
                                    'amount_currency'])[0]}
        used_context = self._build_contexts(data)
        data['form']['used_context'] = dict(
            used_context, lang=get_lang(self.env).code)

        report_env = self.env['report.base_accounting_kit.report_partnerledger']
        report_vals = report_env.with_context(
            data['form'].get('used_context'))._get_report_values(
            data.get('ids', []), data=data)
        partners = report_vals['docs']
        lines_fn = report_vals['lines']
        sum_partner_fn = report_vals['sum_partner']

        partner_rows = []
        for partner in partners:
            p_lines = lines_fn(data, partner)
            p_debit = sum_partner_fn(data, partner, 'debit')
            p_credit = sum_partner_fn(data, partner, 'credit')
            p_balance = sum_partner_fn(data, partner, 'debit - credit')
            partner_rows.append({
                'partner': partner,
                'debit': p_debit,
                'credit': p_credit,
                'balance': p_balance,
                'lines': p_lines,
            })

        partner_rows.sort(key=lambda p: float(p['balance'] or 0.0),
                          reverse=True)

        journal_codes = [j.code for j in self.env['account.journal'].browse(
            data['form'].get('journal_ids') or [])]
        return self._build_xlsx(data['form'], partner_rows, journal_codes)

    def _build_xlsx(self, data_form, partner_rows, journal_codes):
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        helper = self.env['hilatech.xlsx.report.helper']
        formats = helper._xlsx_get_formats(workbook)

        report_name = _('Partner Ledger')
        sheet = workbook.add_worksheet(report_name[:31])
        sheet.hide_gridlines(2)

        widths = [12, 10, 22, 28, 20, 14, 14, 16]
        for col, w in enumerate(widths):
            sheet.set_column(col, col, w)
        ncols = len(widths)

        info = helper._xlsx_standard_info_pairs(data_form, journal_codes)
        result_sel = data_form.get('result_selection')
        if result_sel:
            label_map = {
                'customer': _('Receivable Accounts'),
                'supplier': _('Payable Accounts'),
                'customer_supplier': _('Receivable and Payable Accounts'),
            }
            info.insert(1, (_("Partner's"), label_map.get(result_sel, '')))
        row = helper._xlsx_write_header(workbook, sheet, formats,
                                         report_name, info, ncols)

        headers = [_('Date'), _('JRNL'), _('Account'), _('Ref'),
                   _('Move'), _('Debit'), _('Credit'), _('Progress')]
        row = helper._xlsx_write_table_header(sheet, row, headers, formats)

        for pr in partner_rows:
            partner = pr['partner']
            sheet.merge_range(row, 0, row, 4, partner.display_name or '',
                              formats['level_1'])
            sheet.write_number(row, 5, float(pr['debit'] or 0.0),
                               formats['level_1_num'])
            sheet.write_number(row, 6, float(pr['credit'] or 0.0),
                               formats['level_1_num'])
            sheet.write_number(row, 7, float(pr['balance'] or 0.0),
                               formats['level_1_num'])
            row += 1

            for line in pr['lines'] or []:
                sheet.write(row, 0, str(line.get('date') or ''),
                            formats['detail'])
                sheet.write(row, 1, line.get('code') or '', formats['detail'])
                sheet.write(row, 2, line.get('a_name') or '', formats['detail'])
                sheet.write(row, 3, line.get('ref') or '', formats['detail'])
                sheet.write(row, 4, line.get('displayed_name') or '',
                            formats['detail'])
                sheet.write_number(row, 5, float(line.get('debit') or 0.0),
                                   formats['detail_num'])
                sheet.write_number(row, 6, float(line.get('credit') or 0.0),
                                   formats['detail_num'])
                sheet.write_number(row, 7, float(line.get('progress') or 0.0),
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
