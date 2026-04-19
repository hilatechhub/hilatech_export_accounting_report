# -*- coding: utf-8 -*-
import io
from dateutil.relativedelta import relativedelta
from odoo import models, _
from odoo.exceptions import UserError
from odoo.tools.misc import get_lang

try:
    import xlsxwriter
except ImportError:
    xlsxwriter = None


class AgedPartnerBalanceXlsx(models.TransientModel):
    """Export XLSX pour le wizard Aged Partner Balance."""

    _inherit = 'account.aged.trial.balance'

    def action_export_xlsx(self):
        self.ensure_one()
        if xlsxwriter is None:
            raise UserError(_("La bibliothèque Python 'xlsxwriter' est requise."))

        if self.period_length <= 0:
            raise UserError(_('You must set a period length greater than 0.'))
        if not self.date_from:
            raise UserError(_('You must set a start date.'))

        data = {'ids': self.env.context.get('active_ids', []),
                'model': self.env.context.get('active_model', 'ir.ui.menu'),
                'form': self.read(['date_from', 'date_to', 'journal_ids',
                                    'target_move', 'company_id',
                                    'result_selection', 'period_length'])[0]}
        used_context = self._build_contexts(data)
        data['form']['used_context'] = dict(
            used_context, lang=get_lang(self.env).code)

        period_length = data['form']['period_length']
        start = data['form']['date_from']
        period_labels = {}
        for i in range(5)[::-1]:
            stop = start - relativedelta(days=period_length - 1)
            period_labels[str(i)] = {
                'name': (f"{(5 - (i + 1)) * period_length}-"
                         f"{(5 - i) * period_length}" if i != 0
                         else f"+{4 * period_length}"),
            }
            start = stop - relativedelta(days=1)

        if data['form']['result_selection'] == 'customer':
            account_type = ['asset_receivable']
        elif data['form']['result_selection'] == 'supplier':
            account_type = ['liability_payable']
        else:
            account_type = ['liability_payable', 'asset_receivable']

        report_env = self.env[
            'report.base_accounting_kit.report_agedpartnerbalance']
        move_lines, totals, _dummy = report_env.with_context(
            data['form'].get('used_context'))._get_partner_move_lines(
            account_type, str(data['form']['date_from']),
            data['form']['target_move'], period_length)

        move_lines.sort(key=lambda p: float(p.get('total') or 0.0),
                        reverse=True)

        journal_codes = [j.code for j in self.env['account.journal'].browse(
            data['form'].get('journal_ids') or [])]
        return self._build_xlsx(data['form'], move_lines, totals,
                                 period_labels, journal_codes)

    def _build_xlsx(self, data_form, partners, totals, period_labels,
                    journal_codes):
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        helper = self.env['hilatech.xlsx.report.helper']
        formats = helper._xlsx_get_formats(workbook)

        report_name = _('Aged Partner Balance')
        sheet = workbook.add_worksheet(report_name[:31])
        sheet.hide_gridlines(2)

        sheet.set_column(0, 0, 35)
        sheet.set_column(1, 7, 16)
        ncols = 8

        info = helper._xlsx_standard_info_pairs(data_form, journal_codes)
        info.insert(1, (_('Period Length (days)'),
                        str(data_form.get('period_length'))))
        result_sel = data_form.get('result_selection')
        label_map = {
            'customer': _('Receivable Accounts'),
            'supplier': _('Payable Accounts'),
            'customer_supplier': _('Receivable and Payable Accounts'),
        }
        info.insert(2, (_("Partner's"), label_map.get(result_sel, '')))
        row = helper._xlsx_write_header(workbook, sheet, formats,
                                         report_name, info, ncols)

        headers = [
            _('Partner'), _('Not due'),
            period_labels['4']['name'], period_labels['3']['name'],
            period_labels['2']['name'], period_labels['1']['name'],
            period_labels['0']['name'], _('Total'),
        ]
        row = helper._xlsx_write_table_header(sheet, row, headers, formats)

        sheet.write(row, 0, _('Account Total'), formats['level_1'])
        sheet.write_number(row, 1, float(totals[6] or 0.0),
                           formats['level_1_num'])
        for idx, col in enumerate([4, 3, 2, 1, 0]):
            sheet.write_number(row, 2 + idx, float(totals[col] or 0.0),
                               formats['level_1_num'])
        sheet.write_number(row, 7, float(totals[5] or 0.0),
                           formats['level_1_num'])
        row += 1

        for p in partners:
            sheet.write(row, 0, p.get('name') or '', formats['detail'])
            sheet.write_number(row, 1, float(p.get('direction') or 0.0),
                               formats['detail_num'])
            for idx, key in enumerate(['4', '3', '2', '1', '0']):
                sheet.write_number(row, 2 + idx, float(p.get(key) or 0.0),
                                   formats['detail_num'])
            sheet.write_number(row, 7, float(p.get('total') or 0.0),
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
