# -*- coding: utf-8 -*-
import base64
import io
from odoo import fields, models, _


class XlsxReportHelper(models.AbstractModel):
    """Helper abstrait partagé par tous les wizards d'export Excel.

    Fournit:
    - Des formats xlsxwriter réutilisables (titres, en-têtes, monétaire,
      niveaux hiérarchiques).
    - Un écriteau de bloc en-tête standard (logo société, nom, période,
      filtres) placé en haut de chaque feuille.
    - Une méthode utilitaire pour créer la pièce jointe et retourner
      l'action de téléchargement.
    """

    _name = 'hilatech.xlsx.report.helper'
    _description = "Helper d'export XLSX pour rapports comptables"

    def _xlsx_get_formats(self, workbook):
        """Retourne un dict des formats standards utilisés par les exports."""
        company = self.env.company
        currency_symbol = company.currency_id.symbol or ''
        currency_pos = company.currency_id.position or 'after'
        if currency_pos == 'before':
            money_fmt = f'{currency_symbol} #,##0.00;[Red]-{currency_symbol} #,##0.00'
        else:
            money_fmt = f'#,##0.00 "{currency_symbol}";[Red]-#,##0.00 "{currency_symbol}"'

        return {
            'title': workbook.add_format({
                'bold': True, 'font_size': 16,
                'align': 'center', 'valign': 'vcenter',
                'font_color': '#1F4E78',
            }),
            'company': workbook.add_format({
                'bold': True, 'font_size': 12,
                'align': 'center', 'font_color': '#1F4E78',
            }),
            'info_label': workbook.add_format({
                'bold': True, 'font_size': 10,
                'align': 'left', 'bg_color': '#F2F2F2',
                'border': 1, 'border_color': '#BFBFBF',
            }),
            'info_value': workbook.add_format({
                'font_size': 10, 'align': 'left',
                'border': 1, 'border_color': '#BFBFBF',
            }),
            'table_header': workbook.add_format({
                'bold': True, 'font_size': 11,
                'align': 'center', 'valign': 'vcenter',
                'bg_color': '#1F4E78', 'font_color': 'white',
                'border': 1,
            }),
            'level_1': workbook.add_format({
                'bold': True, 'font_size': 11,
                'bg_color': '#D9E1F2', 'border': 1,
            }),
            'level_1_num': workbook.add_format({
                'bold': True, 'font_size': 11,
                'bg_color': '#D9E1F2', 'border': 1,
                'num_format': money_fmt, 'align': 'right',
            }),
            'level_2': workbook.add_format({
                'bold': True, 'font_size': 10,
                'bg_color': '#E7EEF7', 'border': 1,
            }),
            'level_2_num': workbook.add_format({
                'bold': True, 'font_size': 10,
                'bg_color': '#E7EEF7', 'border': 1,
                'num_format': money_fmt, 'align': 'right',
            }),
            'level_3': workbook.add_format({
                'bold': True, 'font_size': 10,
                'border': 1,
            }),
            'level_3_num': workbook.add_format({
                'bold': True, 'font_size': 10,
                'border': 1,
                'num_format': money_fmt, 'align': 'right',
            }),
            'detail': workbook.add_format({
                'font_size': 10, 'border': 1,
            }),
            'detail_num': workbook.add_format({
                'font_size': 10, 'border': 1,
                'num_format': money_fmt, 'align': 'right',
            }),
            'footer': workbook.add_format({
                'italic': True, 'font_size': 9,
                'font_color': '#808080', 'align': 'right',
            }),
        }

    def _xlsx_write_header(self, workbook, sheet, formats, title,
                            info_pairs, ncols):
        """Écrit le bloc en-tête standard sur les premières lignes.

        - Ligne 0: logo (si disponible) à gauche, titre centré
        - Ligne 2: nom de la société (centré)
        - Ligne 4+: tableau clé/valeur d'informations (period, filtres, etc.)

        Retourne l'index de la prochaine ligne libre.
        """
        company = self.env.company
        last_col = ncols - 1
        sheet.set_row(0, 30)

        if company.logo:
            try:
                image_data = io.BytesIO(base64.b64decode(company.logo))
                sheet.insert_image(
                    0, 0, 'logo.png',
                    {'image_data': image_data,
                     'x_scale': 0.25, 'y_scale': 0.25,
                     'x_offset': 5, 'y_offset': 5},
                )
            except Exception:
                pass

        if last_col >= 1:
            sheet.merge_range(0, 1, 0, last_col, title, formats['title'])
        else:
            sheet.write(0, 0, title, formats['title'])

        if last_col >= 1:
            sheet.merge_range(2, 0, 2, last_col,
                              company.display_name or '', formats['company'])
        else:
            sheet.write(2, 0, company.display_name or '', formats['company'])

        row = 4
        for label, value in info_pairs:
            sheet.write(row, 0, label, formats['info_label'])
            val = value if value is not None else ''
            if last_col >= 2:
                sheet.merge_range(row, 1, row, last_col, val,
                                  formats['info_value'])
            else:
                sheet.write(row, 1, val, formats['info_value'])
            row += 1
        return row + 1

    def _xlsx_standard_info_pairs(self, data_form, journal_codes=None):
        """Construit la liste [(label, value)] standard pour le bloc en-tête.

        `data_form` est typiquement `data['form']` issu du wizard.
        `journal_codes` est une liste de codes journaux à inclure (optionnel).
        """
        target = data_form.get('target_move')
        target_label = (_('All Posted Entries') if target == 'posted'
                        else _('All Entries') if target == 'all' else '-')
        pairs = [
            (_('Target Moves'), target_label),
            (_('Start Date'),
             str(data_form.get('date_from')) if data_form.get('date_from') else '-'),
            (_('End Date'),
             str(data_form.get('date_to')) if data_form.get('date_to') else '-'),
        ]
        if journal_codes:
            pairs.append((_('Journals'), ', '.join(journal_codes)))
        pairs.append((_('Printed on'),
                      fields.Date.context_today(self).isoformat()))
        return pairs

    def _xlsx_write_table_header(self, sheet, row, headers, formats):
        """Écrit une ligne d'en-tête de tableau et retourne la ligne suivante."""
        for col, header in enumerate(headers):
            sheet.write(row, col, header, formats['table_header'])
        return row + 1

    def _xlsx_build_filename(self, report_name, data_form=None):
        """Construit un nom de fichier safe avec suffixe de période."""
        safe = report_name.replace(' ', '_').replace('/', '_').replace('\\', '_')
        date_suffix = ''
        if data_form:
            df = data_form.get('date_from')
            dt = data_form.get('date_to')
            if df or dt:
                date_suffix = f'_{df or ""}_{dt or ""}'
        return f'{safe}{date_suffix}.xlsx'

    def _xlsx_build_attachment_action(self, filename, xlsx_bytes):
        """Crée une ir.attachment pour le fichier xlsx et retourne l'action
        de téléchargement."""
        attachment = self.env['ir.attachment'].create({
            'name': filename,
            'type': 'binary',
            'datas': base64.b64encode(xlsx_bytes),
            'res_model': self._name,
            'mimetype':
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        })
        return {
            'type': 'ir.actions.act_url',
            'url': f'/web/content/{attachment.id}?download=true',
            'target': 'self',
        }
