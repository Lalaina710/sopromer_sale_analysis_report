# Copyright 2026 SOPROMER
# License LGPL-3.0 or later (https://www.gnu.org/licenses/lgpl).
"""Helper AbstractModel pour generer le XLSX Analyse Ventes (format Sage).

Encapsule la generation xlsxwriter dans un AbstractModel pour permettre
l'override par des modules tiers (ex : ajouter colonnes, palette SOPROMER).

Structure du fichier (matche Sage Ventes Sages.xlsx) :
    - Sheet "Ventes"
    - Ligne 1 : titre rapport (merge sur 19 colonnes)
    - Ligne 2 : header 19 colonnes (gras, fond gris)
    - Lignes 3+ : donnees
    - Derniere ligne : totaux sur colonnes numeriques
    - Freeze panes ligne 2
    - Format dates DD/MM/YYYY, montants 2 decimales, qte 3 decimales
"""
import io
import logging

import xlsxwriter

from odoo import _, models

_logger = logging.getLogger(__name__)


class SaleAnalysisReportXlsx(models.AbstractModel):
    _name = 'report.sale.analysis.xlsx'
    _description = 'Generateur XLSX Rapport Analyse Ventes (Sage)'

    # Definition des 19 colonnes : (cle_dict, libelle, type_format, largeur)
    # type_format : 'date' | 'text' | 'int' | 'num' | 'qty'
    COLUMNS = [
        ('date_vente',             "Date Vente",                    'date', 12),
        ('nb_documents',           "Nbre de documents",             'int',  10),
        ('num_piece',              "N° Pièce",            'text', 18),
        ('type_document',          "Type Document",                 'text', 28),
        ('reference',              "Référence",           'text', 18),
        ('document_en_cours',      "Document en cours",             'text', 14),
        ('ca_ht',                  "Chiffre d'affaires HT",         'num',  16),
        ('ca_ttc',                 "Chiffre d'affaires TTC",        'num',  16),
        ('qte_vendue',             "Qté Vendues",              'qty',  12),
        ('prix_revient_total',     "Prix Revient Total",            'num',  16),
        ('marge',                  "Marge",                         'num',  16),
        ('ref_article',            "Référence Article",   'text', 16),
        ('designation',            "Désignation Article",      'text', 38),
        ('code_famille',           "Code Famille",                  'text', 22),
        ('intitule_famille',       "Intitulé Famille",         'text', 22),
        ('num_compte_client',      "N° Compte Client",         'text', 16),
        ('intitule_client',        "Intitulé Client",          'text', 32),
        ('classement_client',      "Classement Client",             'text', 20),
        ('categ_tarifaire_client', "Catégorie Tarifaire Client", 'text', 22),
    ]

    # Index des colonnes a totaliser (CA HT, CA TTC, Qte, PR, Marge)
    TOTAL_COLS = {
        6: 'total_ca_ht',
        7: 'total_ca_ttc',
        8: 'total_qte',
        9: 'total_prix_revient',
        10: 'total_marge',
    }

    def generate(self, data):
        """Genere le XLSX et retourne les bytes.

        Args:
            data: dict produit par wizard._get_report_data()

        Returns:
            bytes: contenu du fichier xlsx
        """
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {
            'in_memory': True,
            'default_date_format': 'dd/mm/yyyy',
        })
        sheet = workbook.add_worksheet('Ventes')

        formats = self._build_formats(workbook)
        self._write_title(sheet, formats, data)
        self._write_header(sheet, formats)
        last_row = self._write_data_rows(sheet, formats, data['rows'])
        self._write_totals(sheet, formats, data['totals'], last_row)
        self._set_column_widths(sheet)

        # Freeze sur la ligne d'entete (ligne 2 = index 2 en xlsxwriter)
        sheet.freeze_panes(2, 0)
        # Autofilter sur l'entete + zone de donnees
        if data['rows']:
            sheet.autofilter(1, 0, last_row, len(self.COLUMNS) - 1)

        workbook.close()
        return output.getvalue()

    # -------------------------------------------------------------------------
    # Formats
    # -------------------------------------------------------------------------
    def _build_formats(self, workbook):
        """Cree les formats xlsxwriter reutilisables."""
        return {
            'title': workbook.add_format({
                'bold': True,
                'font_size': 14,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#1F4E79',
                'font_color': 'white',
                'border': 1,
            }),
            'header': workbook.add_format({
                'bold': True,
                'font_size': 10,
                'bg_color': '#D3D3D3',
                'font_color': 'black',
                'border': 1,
                'align': 'center',
                'valign': 'vcenter',
                'text_wrap': True,
            }),
            'cell_text': workbook.add_format({
                'border': 1,
                'font_size': 10,
            }),
            'cell_date': workbook.add_format({
                'num_format': 'dd/mm/yyyy',
                'border': 1,
                'font_size': 10,
                'align': 'center',
            }),
            'cell_int': workbook.add_format({
                'num_format': '#,##0',
                'border': 1,
                'font_size': 10,
                'align': 'right',
            }),
            'cell_num': workbook.add_format({
                'num_format': '#,##0.00',
                'border': 1,
                'font_size': 10,
                'align': 'right',
            }),
            'cell_num_neg': workbook.add_format({
                'num_format': '#,##0.00',
                'border': 1,
                'font_size': 10,
                'align': 'right',
                'font_color': '#C00000',
            }),
            'cell_qty': workbook.add_format({
                'num_format': '#,##0.000',
                'border': 1,
                'font_size': 10,
                'align': 'right',
            }),
            'total_label': workbook.add_format({
                'bold': True,
                'bg_color': '#1F4E79',
                'font_color': 'white',
                'border': 1,
                'font_size': 10,
                'align': 'right',
            }),
            'total_num': workbook.add_format({
                'bold': True,
                'bg_color': '#1F4E79',
                'font_color': 'white',
                'num_format': '#,##0.00',
                'border': 1,
                'font_size': 10,
                'align': 'right',
            }),
            'total_qty': workbook.add_format({
                'bold': True,
                'bg_color': '#1F4E79',
                'font_color': 'white',
                'num_format': '#,##0.000',
                'border': 1,
                'font_size': 10,
                'align': 'right',
            }),
            'total_int': workbook.add_format({
                'bold': True,
                'bg_color': '#1F4E79',
                'font_color': 'white',
                'num_format': '#,##0',
                'border': 1,
                'font_size': 10,
                'align': 'right',
            }),
        }

    # -------------------------------------------------------------------------
    # Sections
    # -------------------------------------------------------------------------
    def _write_title(self, sheet, formats, data):
        """Ligne 1 : titre fusionne + sous-titre."""
        ncols = len(self.COLUMNS)
        title = _(
            "Rapport Analyse Ventes (Format Sage) - %(company)s - "
            "Periode du %(from)s au %(to)s - %(filters)s"
        ) % {
            'company': data['company'].name,
            'from': data['date_from'].strftime('%d/%m/%Y') if data['date_from'] else '',
            'to': data['date_to'].strftime('%d/%m/%Y') if data['date_to'] else '',
            'filters': data['filters_summary'],
        }
        sheet.set_row(0, 28)
        sheet.merge_range(0, 0, 0, ncols - 1, title, formats['title'])

    def _write_header(self, sheet, formats):
        """Ligne 2 : en-tete des 19 colonnes."""
        sheet.set_row(1, 32)
        for col, (_key, label, _ftype, _w) in enumerate(self.COLUMNS):
            sheet.write(1, col, label, formats['header'])

    def _write_data_rows(self, sheet, formats, rows):
        """Ecrit les lignes de donnees a partir de la ligne 3 (index 2).

        Returns:
            int: index de la derniere ligne ecrite (0-based)
        """
        if not rows:
            return 1
        row_idx = 2
        for r in rows:
            for col, (key, _label, ftype, _w) in enumerate(self.COLUMNS):
                self._write_cell(sheet, row_idx, col, r.get(key, ''), ftype, formats)
            row_idx += 1
        return row_idx - 1

    def _write_cell(self, sheet, row, col, value, ftype, formats):
        """Ecrit une cellule en choisissant le bon format selon le type."""
        if ftype == 'date':
            if value:
                sheet.write_datetime(row, col, value, formats['cell_date'])
            else:
                sheet.write_blank(row, col, '', formats['cell_date'])
        elif ftype == 'int':
            sheet.write_number(row, col, int(value or 0), formats['cell_int'])
        elif ftype == 'num':
            num = float(value or 0.0)
            fmt = formats['cell_num_neg'] if num < 0 else formats['cell_num']
            sheet.write_number(row, col, num, fmt)
        elif ftype == 'qty':
            sheet.write_number(row, col, float(value or 0.0), formats['cell_qty'])
        else:
            # text (defaut)
            sheet.write_string(row, col, str(value) if value else '', formats['cell_text'])

    def _write_totals(self, sheet, formats, totals, last_row):
        """Ecrit la ligne TOTAUX sous les donnees."""
        row = last_row + 1
        ncols = len(self.COLUMNS)

        # Label TOTAUX merge sur les 6 premieres colonnes (avant CA HT)
        sheet.merge_range(row, 0, row, 5, _("TOTAUX"), formats['total_label'])

        # Colonnes numeriques totalisees
        for col_idx, total_key in self.TOTAL_COLS.items():
            value = totals.get(total_key, 0.0)
            if col_idx == 8:  # qte_vendue (3 decimales)
                sheet.write_number(row, col_idx, value, formats['total_qty'])
            else:
                sheet.write_number(row, col_idx, value, formats['total_num'])

        # Cellules vides restantes (apres marge) -> on remplit avec format total
        for col in range(11, ncols):
            sheet.write_blank(row, col, '', formats['total_label'])

        # Ligne resume sous les totaux : nb_lignes / nb_documents
        summary_row = row + 1
        summary_text = _(
            "%(lines)d lignes - %(docs)d documents - Genere le %(printed)s"
        ) % {
            'lines': totals.get('nb_lignes', 0),
            'docs': totals.get('nb_documents', 0),
            'printed': '',  # rempli par le wizard caller si besoin
        }
        # On evite la double-ecriture du print_date ici - pas critique
        sheet.merge_range(summary_row, 0, summary_row, ncols - 1, summary_text,
                          formats['cell_text'])

    def _set_column_widths(self, sheet):
        """Applique la largeur de chaque colonne."""
        for col, (_key, _label, _ftype, width) in enumerate(self.COLUMNS):
            sheet.set_column(col, col, width)
