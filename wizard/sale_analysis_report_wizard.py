# Copyright 2026 SOPROMER
# License LGPL-3.0 or later (https://www.gnu.org/licenses/lgpl).
"""Wizard rapport analyse ventes (format Sage) - Excel only.

Genere un rapport Excel reproduisant le rapport Sage "Ventes" (15k+ lignes)
pour analyse TCD cote client.

Granularite : 1 ligne par account.move.line (lignes de type 'product').
Source     : factures clients (out_invoice + out_refund) HORS POS.
Devise     : MGA (mono-devise SOPROMER).

Architecture :
    - _build_domain()           : domain ORM avec exclusion POS
    - _get_pos_move_ids()       : factures issues du POS (a exclure)
    - _get_report_data()        : extraction + transformation
    - _row_from_line(line)      : mapping 1 line -> dict 19 colonnes
    - _compute_totals(rows)     : totaux globaux
    - action_generate_report()  : genere XLSX et retourne URL download
"""
import base64
import logging
from datetime import date

from odoo import _, api, fields, models
from odoo.exceptions import UserError, ValidationError

_logger = logging.getLogger(__name__)


class SaleAnalysisReportWizard(models.TransientModel):
    _name = 'sale.analysis.report.wizard'
    _description = 'Wizard Rapport Analyse Ventes (Format Sage)'

    # =========================================================================
    # FILTRES PERIODE (obligatoires)
    # =========================================================================
    date_from = fields.Date(
        string="Date debut",
        required=True,
        default=lambda self: date.today().replace(day=1),
        help="Date de debut (incluse). Filtre sur account.move.invoice_date.",
    )
    date_to = fields.Date(
        string="Date fin",
        required=True,
        default=fields.Date.context_today,
        help="Date de fin (incluse).",
    )

    # =========================================================================
    # FILTRES OPTIONNELS
    # =========================================================================
    partner_ids = fields.Many2many(
        'res.partner',
        'sale_analysis_wiz_partner_rel',
        'wizard_id', 'partner_id',
        string="Clients",
        domain="[('customer_rank', '>', 0)]",
        help="Filtre clients. Vide = tous les clients.",
    )
    pricelist_ids = fields.Many2many(
        'product.pricelist',
        'sale_analysis_wiz_pricelist_rel',
        'wizard_id', 'pricelist_id',
        string="Categorie tarifaire client",
        help="Filtre par liste de prix client (proxy 'Categorie Tarifaire' Sage : "
             "Detail, Grossiste, etc.). Vide = toutes categories.",
    )
    category_ids = fields.Many2many(
        'product.category',
        'sale_analysis_wiz_categ_rel',
        'wizard_id', 'category_id',
        string="Familles articles",
        help="Filtre par famille de produits (product.category). "
             "Inclut les sous-categories (child_of). Vide = toutes familles.",
    )

    # =========================================================================
    # MULTI-COMPANY
    # =========================================================================
    company_id = fields.Many2one(
        'res.company',
        string="Societe",
        required=True,
        default=lambda self: self.env.company,
        help="Societe pour laquelle generer le rapport.",
    )

    # =========================================================================
    # OPTIONS
    # =========================================================================
    include_draft = fields.Boolean(
        string="Inclure brouillons",
        default=False,
        help="Inclure les factures en brouillon (state='draft'). "
             "Par defaut, seules les factures comptabilisees (posted) sont incluses.",
    )

    # =========================================================================
    # OUTPUT (rempli par action_generate_report)
    # =========================================================================
    report_xlsx_file = fields.Binary(string="Fichier Excel", readonly=True)
    report_xlsx_filename = fields.Char(string="Nom fichier Excel", readonly=True)

    # =========================================================================
    # CONSTRAINTS
    # =========================================================================
    @api.constrains('date_from', 'date_to')
    def _check_dates(self):
        for record in self:
            if record.date_from and record.date_to and record.date_from > record.date_to:
                raise ValidationError(
                    _("La date de debut doit etre anterieure ou egale a la date de fin.")
                )

    # =========================================================================
    # DATA EXTRACTION
    # =========================================================================
    def _is_pos_installed(self):
        """Verifie si le module point_of_sale est installe."""
        module = self.env['ir.module.module'].sudo().search([
            ('name', '=', 'point_of_sale'),
            ('state', '=', 'installed'),
        ], limit=1)
        return bool(module)

    def _get_pos_move_ids(self):
        """Retourne la liste des account.move.id lies a une commande POS.

        Ces factures sont exclues du rapport (rapport ventes hors POS).

        Returns:
            list[int]: IDs des account.move issus de POS, ou [] si POS pas installe.
        """
        if not self._is_pos_installed():
            return []
        # sudo() : la lecture des pos.order traverse la securite multi-shop
        # qui peut limiter la visibilite a la session courante. On veut TOUTES
        # les factures POS pour les exclure correctement.
        pos_orders = self.env['pos.order'].sudo().search([
            ('account_move', '!=', False),
        ])
        return pos_orders.mapped('account_move').ids

    def _build_domain(self):
        """Construit le domain ORM pour rechercher les lignes de facture.

        Filtres applique :
          - Etat : posted (+ draft si include_draft)
          - Type : out_invoice + out_refund (factures clients)
          - Periode : invoice_date dans [date_from, date_to]
          - Societe : company_id
          - Lignes produit uniquement (display_type = 'product')
          - Hors POS (move_id not in pos_move_ids)
          - Clients / pricelists / familles (optionnels)

        Returns:
            list: domain Odoo standard
        """
        self.ensure_one()
        states = ['posted', 'draft'] if self.include_draft else ['posted']
        domain = [
            ('parent_state', 'in', states),
            ('move_id.move_type', 'in', ('out_invoice', 'out_refund')),
            ('move_id.invoice_date', '>=', self.date_from),
            ('move_id.invoice_date', '<=', self.date_to),
            ('move_id.company_id', '=', self.company_id.id),
            ('display_type', '=', 'product'),
        ]

        # Exclusion POS
        pos_move_ids = self._get_pos_move_ids()
        if pos_move_ids:
            domain.append(('move_id', 'not in', pos_move_ids))

        # Filtres optionnels
        if self.partner_ids:
            domain.append(('move_id.partner_id', 'in', self.partner_ids.ids))
        if self.category_ids:
            domain.append(('product_id.categ_id', 'child_of', self.category_ids.ids))
        if self.pricelist_ids:
            partners = self.env['res.partner'].search([
                ('property_product_pricelist', 'in', self.pricelist_ids.ids),
            ])
            # Force [0] pour eviter le domain vide qui matcherait tout
            domain.append(('move_id.partner_id', 'in', partners.ids or [0]))

        return domain

    def _get_report_data(self):
        """Centralise extraction + transformation des donnees du rapport.

        Returns:
            dict: {
                'company': res.company,
                'date_from': date,
                'date_to': date,
                'rows': list[dict],   # 1 dict par ligne de facture
                'totals': dict,       # totaux globaux
                'filters_summary': str,
                'print_date': datetime,
            }
        """
        self.ensure_one()
        domain = self._build_domain()

        # Recherche + tri stable (date, piece, id)
        AccountMoveLine = self.env['account.move.line']
        lines = AccountMoveLine.search(domain, order='invoice_date, move_id, id')

        # Prefetch pour eviter N+1 queries (essentiel sur 15k lignes)
        # Le simple acces a un champ many2one declenche le prefetch automatique
        # sur tout le recordset, donc on amorce explicitement les acces.
        lines.mapped('move_id.partner_id.property_product_pricelist.name')
        lines.mapped('product_id.categ_id.complete_name')
        lines.mapped('product_id.standard_price')

        _logger.info(
            "Rapport analyse ventes : %d lignes a traiter (periode %s - %s, societe %s)",
            len(lines), self.date_from, self.date_to, self.company_id.name,
        )

        rows = [self._row_from_line(line) for line in lines]
        totals = self._compute_totals(rows)

        return {
            'company': self.company_id,
            'date_from': self.date_from,
            'date_to': self.date_to,
            'rows': rows,
            'totals': totals,
            'filters_summary': self._build_filters_summary(),
            'print_date': fields.Datetime.now(),
        }

    def _row_from_line(self, line):
        """Transforme un account.move.line en dict ligne rapport (19 colonnes Sage).

        Mapping Sage <-> Odoo :
            1.  Date Vente            : invoice_date (fallback date)
            2.  Nbre documents        : 1 (constante)
            3.  N Piece               : move.name
            4.  Type Document         : "Facture comptabilisee" /
                                        "Facture de retour comptabilisee"
                                        + concat etat si draft
            5.  Reference             : move.ref
            6.  Document en cours     : OUI (draft) / NON (posted)
            7.  CA HT                 : price_subtotal (signe inverse si refund)
            8.  CA TTC                : price_total (signe inverse si refund)
            9.  Qte vendue            : quantity (signe inverse si refund)
            10. Prix Revient Total    : standard_price * quantity (CMUP x qty)
            11. Marge                 : ca_ht - prix_revient_total (vraie marge)
            12. Reference Article     : product.default_code
            13. Designation Article   : product.display_name (fallback line.name)
            14. Code Famille          : product.categ_id.complete_name
            15. Intitule Famille      : product.categ_id.name
            16. N Compte Client       : partner.ref (fallback partner.id)
            17. Intitule Client       : partner.name
            18. Classement Client     : partner.property_product_pricelist.name
            19. Categorie Tarifaire   : partner.property_product_pricelist.name
                                        (meme valeur que col 18, expose 2x dans Sage)

        Returns:
            dict: ligne du rapport prete pour XLSX
        """
        move = line.move_id
        product = line.product_id
        partner = move.partner_id
        is_refund = move.move_type == 'out_refund'
        sign = -1 if is_refund else 1

        # Type document : libelle Sage
        if is_refund:
            base_label = _("Facture de retour")
        else:
            base_label = _("Facture")
        if move.state == 'draft':
            type_doc = _("%s (brouillon)") % base_label
        else:
            type_doc = _("%s comptabilisee") % base_label

        # CMUP / cout standard
        std_price = product.standard_price if product else 0.0
        qty_signed = line.quantity * sign
        ca_ht = line.price_subtotal * sign
        ca_ttc = line.price_total * sign
        prix_revient_total = std_price * qty_signed
        marge = ca_ht - prix_revient_total

        # Famille produit
        categ = product.categ_id if product else False
        code_famille = categ.complete_name if categ else ''
        intitule_famille = categ.name if categ else ''

        # Categorie tarifaire client (cols 18 et 19 = meme valeur, format Sage)
        pricelist = partner.property_product_pricelist if partner else False
        pricelist_name = pricelist.name if pricelist else ''

        # N compte client : ref si dispo, sinon ID Odoo formate
        if partner.ref:
            num_compte = partner.ref
        elif partner:
            num_compte = "C%07d" % partner.id
        else:
            num_compte = ''

        # Designation : display_name produit ou fallback label de ligne
        if product:
            designation = product.display_name
        else:
            designation = line.name or ''

        return {
            'date_vente': move.invoice_date or move.date,
            'nb_documents': 1,
            'num_piece': move.name or '',
            'type_document': type_doc,
            'reference': move.ref or '',
            'document_en_cours': _("OUI") if move.state == 'draft' else _("NON"),
            'ca_ht': ca_ht,
            'ca_ttc': ca_ttc,
            'qte_vendue': qty_signed,
            'prix_revient_total': prix_revient_total,
            'marge': marge,
            'ref_article': product.default_code if product else '',
            'designation': designation,
            'code_famille': code_famille,
            'intitule_famille': intitule_famille,
            'num_compte_client': num_compte,
            'intitule_client': partner.name or '',
            'classement_client': pricelist_name,
            'categ_tarifaire_client': pricelist_name,
        }

    @staticmethod
    def _compute_totals(rows):
        """Calcule les totaux globaux du rapport sur colonnes numeriques."""
        return {
            'nb_lignes': len(rows),
            'nb_documents': len({r['num_piece'] for r in rows if r['num_piece']}),
            'total_ca_ht': sum(r['ca_ht'] for r in rows),
            'total_ca_ttc': sum(r['ca_ttc'] for r in rows),
            'total_qte': sum(r['qte_vendue'] for r in rows),
            'total_prix_revient': sum(r['prix_revient_total'] for r in rows),
            'total_marge': sum(r['marge'] for r in rows),
        }

    def _build_filters_summary(self):
        """Construit un libelle resume des filtres actifs (pour entete rapport)."""
        self.ensure_one()
        parts = []
        if self.partner_ids:
            parts.append(_("Clients : %d") % len(self.partner_ids))
        if self.pricelist_ids:
            parts.append(_("Tarifs : %s") % ', '.join(self.pricelist_ids.mapped('name')))
        if self.category_ids:
            parts.append(_("Familles : %d") % len(self.category_ids))
        if self.include_draft:
            parts.append(_("Brouillons inclus"))
        return ' | '.join(parts) if parts else _("Aucun filtre")

    # =========================================================================
    # ACTIONS
    # =========================================================================
    def action_generate_report(self):
        """Genere le rapport Excel et retourne l'URL de telechargement.

        Returns:
            dict: ir.actions.act_url vers /web/content (download direct)
        """
        self.ensure_one()
        _logger.info(
            "Generation rapport analyse ventes XLSX du %s au %s pour societe %s",
            self.date_from, self.date_to, self.company_id.name,
        )

        data = self._get_report_data()
        if not data['rows']:
            raise UserError(
                _("Aucune ligne de facture trouvee pour les criteres selectionnes.")
            )

        XlsxBuilder = self.env['report.sale.analysis.xlsx']
        content = XlsxBuilder.generate(data)
        filename = self._build_filename()

        self.write({
            'report_xlsx_file': base64.b64encode(content),
            'report_xlsx_filename': filename,
        })

        _logger.info(
            "Rapport analyse ventes genere : %s (%d lignes, %d documents)",
            filename, data['totals']['nb_lignes'], data['totals']['nb_documents'],
        )

        return {
            'type': 'ir.actions.act_url',
            'url': (
                f'/web/content/?model=sale.analysis.report.wizard&id={self.id}'
                f'&field=report_xlsx_file&filename_field=report_xlsx_filename'
                f'&download=true'
            ),
            'target': 'self',
        }

    def _build_filename(self):
        """Construit le nom de fichier standardise SOPROMER."""
        return "Ventes_SOPROMER_%s_%s.xlsx" % (self.date_from, self.date_to)
