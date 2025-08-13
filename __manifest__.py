{
    "name": "Sales Weekly Report",
    "summary": """
        Generate a weekly sales report for products by selected categories""",
    "author": "Be OnlyOne",
    "maintainers": ["onlyone-odoo"],
    "website": "https://onlyone.odoo.com/",
    "license": "AGPL-3",
    "category": "Technical Settings",
    "version": "17.0.1.0.0",
    "development_status": "Production/Stable",
    "application": False,
    "installable": True,
    "external_dependencies": {
        "python": [],
        "bin": [],
    },
    "depends": ["account", "product", "l10n_ar", "product_nota_pedido_export"],
    "data": [
        "security/ir.model.access.csv",
        "views/sales_weekly_report_wizard.xml",
    ],
}
