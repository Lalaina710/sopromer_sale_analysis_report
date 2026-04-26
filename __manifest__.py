# Copyright 2026 SOPROMER
# License LGPL-3.0 or later (https://www.gnu.org/licenses/lgpl).
{
    'name': 'Rapport Analyse Ventes (Format Sage)',
    'version': '18.0.1.0.0',
    'category': 'Sales/Reporting',
    'summary': 'Rapport analyse ventes au format Sage - export Excel avec marge reelle, hors POS',
    'description': """
Rapport Analyse Ventes - Format Sage
=====================================
Reproduit le rapport Sage "Ventes" pour analyse TCD cote client.

Caracteristiques :
- Source : factures clients (out_invoice + out_refund), HORS POS
- Granularite : 1 ligne par ligne de facture (account.move.line)
- Filtres : periode, clients, pricelists (categorie tarifaire), familles articles
- 19 colonnes Sage : Date, N piece, Type doc, CA HT/TTC, Qte, Prix Revient,
  MARGE, Article, Client, Categorie, Famille
- Marge reelle : price_subtotal - (standard_price * quantity)
- Devise : MGA (mono-devise SOPROMER)
- Export Excel uniquement (xlsxwriter)
    """,
    'author': 'SOPROMER',
    'website': 'https://github.com/Lalaina710/sopromer_sale_analysis_report',
    'license': 'LGPL-3',
    'depends': [
        'sale_management',
        'account',
        'product',
    ],
    'data': [
        'security/ir.model.access.csv',
        'wizard/sale_analysis_report_wizard_view.xml',
        'views/menu.xml',
    ],
    'external_dependencies': {
        'python': ['xlsxwriter'],
    },
    'installable': True,
    'application': False,
    'auto_install': False,
}
