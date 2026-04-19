# -*- coding: utf-8 -*-
{
    'name': 'Hilatech Export Accounting Reports to Excel',
    'version': '19.0.1.0.0',
    'category': 'Accounting',
    'summary': "Ajoute un bouton 'Export Excel' sur tous les wizards de rapports "
               "comptables du module base_accounting_kit et sur les factures.",
    'description': """
Hilatech Export Accounting Reports
==================================
Ce module ajoute des boutons d'export Excel (XLSX) aux wizards de rapports
comptables (Profit and Loss, Balance Sheet, General Ledger, Trial Balance,
Partner Ledger, Aged Partner Balance, Journal Audit, Cash Flow, Bank Book,
Cash Book, Day Book, Tax Report) ainsi qu'aux factures clients/fournisseurs.

Les exports reproduisent la mise en forme hiérarchique des rapports PDF avec:
- En-tête (société, logo, période, filtres)
- Indentation des comptes enfants
- Totaux en gras
- Formatage monétaire
""",
    'author': 'Olivier Kamdem',
    'company': 'Hilatech',
    'maintainer': 'Hilatech',
    'website': 'https://hilatech.co',
    'support': 'support@hilatech.co',
    'icon': 'static/description/icon.png',
    'images': ['static/description/banner.png'],
    'price': 5.0,
    'currency': 'EUR',
    'depends': [
        'account',
        'base_accounting_kit',
    ],
    'data': [
        'wizard/financial_report_views.xml',
        'wizard/account_report_common_views.xml',
        'wizard/aged_partner_balance_views.xml',
        'wizard/bank_book_views.xml',
        'wizard/cash_book_views.xml',
        'wizard/day_book_views.xml',
        'views/account_move_views.xml',
    ],
    'external_dependencies': {
        'python': ['xlsxwriter'],
    },
    'license': 'LGPL-3',
    'installable': True,
    'auto_install': False,
    'application': False,
}
