# -*- coding: utf-8 -*-
import io
from odoo import models, _
from odoo.exceptions import UserError
from odoo.tools.misc import get_lang

try:
    import xlsxwriter
except ImportError:
    xlsxwriter = None


class JournalAuditXlsx(models.TransientModel):
    """Export XLSX pour le wizard Journal Audit."""

    _inherit = 'account.print.journal'

    def action_export_xlsx(self):
        self.ensure_one()
        if xlsxwriter is None:
            raise UserError(_("La bibliothèque Python 'xlsxwriter' est requise."))

        data = {'ids': self.env.context.get('active_ids', []),
                'model': self.env.context.get('active_model', 'ir.ui.menu'),
                'form': self.read(['date_from', 'date_to', 'journal_ids',
                                    'target_move', 'company_id',
                                    'sort_selection', 'amount_currency'])[0]}
        used_context = self._build_contexts(data)
        data['form']['used_context'] = dict(
            used_context, lang=get_lang(self.env).code)

        report_env = self.env[
            'report.base_accounting_kit.report_journal_audit']
        journals = self.env['account.journal'].browse(
            data['form']['journal_ids'] or [])

        journal_data = []
        for journal in journals:
            lines = report_env.with_context(
                data['form'].get('used_context', {})).lines(
                data['form']['target_move'], journal.id,
                data['form']['sort_selection'], data)
            sum_debit = report_env._sum_debit(data, journal)
            sum_credit = report_env._sum_credit(data, journal)
            taxes = report_env._get_taxes(data, journal)
            journal_data.append({
                'journal': journal,
                'lines': lines,
                'sum_debit': sum_debit,
                'sum_credit': sum_credit,
                'taxes': taxes,
            })

        journal_codes = [j.code for j in journals]
        return self._build_xlsx(data['form'], journal_data, journal_codes)

    def _build_xlsx(self, data_form, journal_data, journal_codes):
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        helper = self.env['hilatech.xlsx.report.helper']
        formats = helper._xlsx_get_formats(workbook)

        report_name = _('Journal Audit')
        sheet = workbook.add_worksheet(report_name[:31])
        sheet.hide_gridlines(2)

        widths = [12, 18, 22, 28, 22, 14, 14]
        for col, w in enumerate(widths):
            sheet.set_column(col, col, w)
        ncols = len(widths)

        info = helper._xlsx_standard_info_pairs(data_form, journal_codes)
        row = helper._xlsx_write_header(workbook, sheet, formats,
                                         report_name, info, ncols)

        headers = [_('Date'), _('Move'), _('Account'), _('Partner'),
                   _('Label'), _('Debit'), _('Credit')]
        row = helper._xlsx_write_table_header(sheet, row, headers, formats)

        for jd in journal_data:
            journal = jd['journal']
            sheet.merge_range(row, 0, row, ncols - 3,
                              f"{journal.code} - {journal.name}",
                              formats['level_1'])
            sheet.write_number(row, 5, float(jd['sum_debit'] or 0.0),
                               formats['level_1_num'])
            sheet.write_number(row, 6, float(jd['sum_credit'] or 0.0),
                               formats['level_1_num'])
            row += 1

            for line in jd['lines']:
                sheet.write(row, 0, str(line.date or ''), formats['detail'])
                sheet.write(row, 1, line.move_id.name or '', formats['detail'])
                sheet.write(row, 2,
                            f"{line.account_id.code or ''} {line.account_id.name or ''}".strip(),
                            formats['detail'])
                sheet.write(row, 3, line.partner_id.name or '',
                            formats['detail'])
                sheet.write(row, 4, line.name or '', formats['detail'])
                sheet.write_number(row, 5, float(line.debit or 0.0),
                                   formats['detail_num'])
                sheet.write_number(row, 6, float(line.credit or 0.0),
                                   formats['detail_num'])
                row += 1

            if jd['taxes']:
                sheet.merge_range(row, 0, row, ncols - 1, _('Taxes'),
                                  formats['level_2'])
                row += 1
                for tax, tax_vals in jd['taxes'].items():
                    sheet.write(row, 0, '', formats['detail'])
                    sheet.merge_range(row, 1, row, 4, tax.name or '',
                                      formats['detail'])
                    sheet.write_number(row, 5,
                                       float(tax_vals.get('base_amount') or 0.0),
                                       formats['detail_num'])
                    sheet.write_number(row, 6,
                                       float(tax_vals.get('tax_amount') or 0.0),
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
