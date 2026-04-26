# sopromer_sale_analysis_report

Rapport Analyse Ventes au format Sage 100 pour SOPROMER (Odoo 18).

Reproduit le rapport Sage "Ventes" utilise comme source pour les TCD cote client.
Export **Excel uniquement** (xlsxwriter), 19 colonnes.

## Caracteristiques

- Source : factures clients (`out_invoice` + `out_refund`), HORS POS
- Granularite : 1 ligne par ligne de facture (`account.move.line`)
- Devise : MGA (mono-devise SOPROMER)
- 19 colonnes Sage : Date, N piece, Type doc, Reference, Etat, CA HT/TTC, Qte,
  Prix Revient, Marge, Article (ref + designation), Famille (code + intitule),
  Client (compte + nom), Classement client, Categorie tarifaire
- **Marge** = `price_subtotal - (standard_price * quantity)` (vraie marge brute)
- Filtres : periode (obligatoire), clients, pricelists, familles articles
- Multi-company

## Installation

```bash
# Sur serveur 45 (test) puis 43 (prod)
docker exec -u odoo odoo-dev /opt/odoo/odoo-bin \
    -c /etc/odoo/odoo.conf -d <db> \
    -i sopromer_sale_analysis_report --stop-after-init --no-http
```

## Utilisation

Menu : **Ventes > Analyse > Rapport Ventes Sage**

1. Choisir la periode (par defaut : 1er du mois courant -> aujourd'hui)
2. (Optionnel) filtrer par clients / pricelists / familles
3. Cliquer "Generer le rapport" : telechargement automatique du fichier
   `Ventes_SOPROMER_<date_from>_<date_to>.xlsx`

## Securite

- Acces : groupes `sales_team.group_sale_salesman` et `account.group_account_invoice`
- Multi-company : filtre automatique sur `company_id`
- Hors POS : exclusion via `pos.order.account_move`

## Developpement

| Fichier | Role |
|---------|------|
| `wizard/sale_analysis_report_wizard.py` | Logique metier (domain, mapping, action) |
| `wizard/sale_analysis_report_wizard_view.xml` | Form view + action wizard |
| `report/sale_analysis_report_xlsx.py` | Generation xlsxwriter (AbstractModel) |
| `views/menu.xml` | Menu sous Ventes > Analyse |
| `security/ir.model.access.csv` | ACL wizard |

## Auteur

SOPROMER - License LGPL-3
