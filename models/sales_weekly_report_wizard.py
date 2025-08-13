# -*- coding: utf-8 -*-
from odoo import models, fields, api, _
from datetime import timedelta
import re  # Para parsear el número de factura


class SalesWeeklyReportWizard(models.TransientModel):
    _name = "sales.weekly.report.wizard"
    _description = "Sales Weekly Report Wizard"

    date_from = fields.Date(
        string="Desde",
        required=True,
        default=lambda self: fields.Date.today() - timedelta(days=7),
    )
    date_to = fields.Date(string="Hasta", required=True, default=fields.Date.today)
    category_ids = fields.Many2many(
        "product.category", string="Categorías de Productos", required=True
    )
    line_ids = fields.One2many(
        "sales.weekly.report.line", "wizard_id", string="Líneas de Reporte"
    )

    def generate_report(self):
        """Generate a weekly sales report for products in selected categories."""
        self.line_ids.unlink()  # Clear previous report lines

        # Fetch all products belonging to selected categories
        product_ids = (
            self.env["product.product"]
            .search([("categ_id", "in", self.category_ids.ids)])
            .ids
        )

        # Search invoice lines for the selected products and date range
        domain = [
            (
                "move_id.move_type",
                "in",
                ("out_invoice", "out_refund"),
            ),  # Incluir refunds para NC
            ("move_id.state", "=", "posted"),
            ("move_id.invoice_date", ">=", self.date_from),
            ("move_id.invoice_date", "<=", self.date_to),
            ("product_id", "in", product_ids),
        ]
        invoice_lines = self.env["account.move.line"].search(domain)

        # Prepare data for report lines using mapped for efficiency
        vals_list = []
        for line in invoice_lines:
            move = line.move_id
            partner = move.partner_id
            # Parsear número de factura: ej. 'FACT-A-0001-00000001' -> punto '0001', numero '00000001'
            match = re.match(r".*-.*-(\d+)-(\d+)", move.name)
            punto_venta = match.group(1) if match else ""
            numero_factura = match.group(2) if match else move.name
            # Tipo factura: FC o NC
            tipo_factura = "FC" if move.move_type == "out_invoice" else "NC"
            # Tipo documento: A/B/C de l10n_latam_document_type_id
            tipo_documento = move.l10n_latam_document_type_id.code or ""
            # Otros campos
            vals = {
                "wizard_id": self.id,
                "invoice_date": move.invoice_date,
                "customer_cuit": partner.vat or "",
                "customer_name": partner.name,
                "invoice_number": move.name,
                "product_name": line.product_id.name,
                "quantity": line.quantity,
                "unit_price": line.price_unit,
                "total_amount": line.price_subtotal,
                "zip_code": partner.zip or "",
                "city": partner.city or "",
                "province": partner.state_id.name or "",
                "tipo_factura": tipo_factura,
                "tipo_documento": tipo_documento,
                "punto_venta": punto_venta,
                "numero_factura": numero_factura,
                "codigo_articulo": line.product_id.default_code or "",
                "rubro": line.product_id.ndp_rubro or "",
                "representante": move.invoice_user_id.name or "",
            }
            vals_list.append(vals)

        # Create report lines
        self.env["sales.weekly.report.line"].create(vals_list)

        # Return action to display the wizard with results
        return {
            "name": _("Reporte Semanal de Ventas"),
            "view_mode": "form",
            "res_model": "sales.weekly.report.wizard",
            "res_id": self.id,
            "type": "ir.actions.act_window",
            "target": "new",
        }

    def export_to_excel(self):
        """Export the report lines to an XLSX file."""
        import io
        import xlsxwriter

        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {"in_memory": True})
        worksheet = workbook.add_worksheet()

        # Cabeceras del Excel
        headers = [
            "NRO.",
            "RAZON SOCIAL",
            "RUBRO",
            "CODIGO ARTICULO",
            "CANT.",
            "$ PRECIO",
            "TIPO",
            "PUNTO",
            "NUMERO",
            "FECHA",
            "CODIGO",
            "LOCALIDAD",
            "PROVINCIA",
            "CUIT",
            "REPRESENTANTE",
            "TIPO DOCUMENTO",  # Agregamos la columna sin título
        ]
        for col_num, header in enumerate(headers):
            worksheet.write(0, col_num, header)

        # Escribir datos de las líneas, con NRO. secuencial
        for row_num, line in enumerate(self.line_ids, start=1):
            worksheet.write(row_num, 0, row_num)  # NRO.
            worksheet.write(row_num, 1, line.customer_name)
            worksheet.write(row_num, 2, line.rubro)
            worksheet.write(row_num, 3, line.codigo_articulo)
            worksheet.write(row_num, 4, line.quantity)
            worksheet.write(row_num, 5, line.unit_price)
            worksheet.write(row_num, 6, line.tipo_factura)
            worksheet.write(row_num, 7, line.punto_venta)
            worksheet.write(row_num, 8, line.numero_factura)
            worksheet.write(row_num, 9, line.invoice_date)
            worksheet.write(row_num, 10, line.zip_code)
            worksheet.write(row_num, 11, line.city)
            worksheet.write(row_num, 12, line.province)
            worksheet.write(row_num, 13, line.customer_cuit)
            worksheet.write(row_num, 14, line.representante)
            worksheet.write(row_num, 15, line.tipo_documento)  # Tipo A/B/C

        workbook.close()
        output.seek(0)
        file_data = output.read()
        output.close()

        # Crear attachment y devolver acción de descarga
        attachment = self.env["ir.attachment"].create(
            {
                "name": "reporte_ventas_semanal.xlsx",
                "type": "binary",
                "datas": file_data.encode("base64"),
                "store_fname": "reporte_ventas_semanal.xlsx",
                "res_model": self._name,
                "res_id": self.id,
                "mimetype": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            }
        )
        return {
            "type": "ir.actions.act_url",
            "url": f"/web/content/?model=ir.attachment&id={attachment.id}&filename_field=name&field=datas&download=true&name={attachment.name}",
            "target": "self",
        }


class SalesWeeklyReportLine(models.TransientModel):
    _name = "sales.weekly.report.line"
    _description = "Línea de Reporte Semanal de Ventas"

    wizard_id = fields.Many2one("sales.weekly.report.wizard", string="Wizard")
    invoice_date = fields.Date(string="Fecha Factura")
    customer_cuit = fields.Char(string="CUIT Cliente")
    customer_name = fields.Char(string="Razón Social")
    invoice_number = fields.Char(string="Nº Factura")
    product_name = fields.Char(string="Producto")
    quantity = fields.Float(string="Cantidad")
    unit_price = fields.Float(string="Precio Unitario")
    total_amount = fields.Float(string="Importe Total")
    zip_code = fields.Char(string="Código Postal")
    city = fields.Char(string="Localidad")
    province = fields.Char(string="Provincia")
    # Campos nuevos para el export
    tipo_factura = fields.Char(string="Tipo (FC/NC)")
    tipo_documento = fields.Char(string="Tipo Documento (A/B/C)")
    punto_venta = fields.Char(string="Punto de Venta")
    numero_factura = fields.Char(string="Número Factura")
    codigo_articulo = fields.Char(string="Código Artículo")
    rubro = fields.Char(string="Rubro")
    representante = fields.Char(string="Representante")
