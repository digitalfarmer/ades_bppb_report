import xlwt
import datetime
import base64
from io import StringIO
from datetime import datetime
from odoo import api, fields, models, _
import platform


style0 = xlwt.easyxf('font: name Times New Roman bold on;align: horiz right;', num_format_str='#,##0.00')
style1 = xlwt.easyxf(
    'font: name Times New Roman bold on; pattern: pattern solid, fore_colour gray25;align: horiz center;',
    num_format_str='#,##0.00')
style2 = xlwt.easyxf('font:height 400,bold True; pattern: pattern solid, fore_colour gray25;',
                     num_format_str='#,##0.00')
style3 = xlwt.easyxf('font:bold True;', num_format_str='#,##0.00')
style4 = xlwt.easyxf('font:bold True;  borders:top double;align: horiz right;', num_format_str='#,##0.00')
style5 = xlwt.easyxf('font: name Times New Roman bold on;align: horiz center;', num_format_str='#,##0')
style6 = xlwt.easyxf('font: name Times New Roman bold on;', num_format_str='#,##0.00')
style7 = xlwt.easyxf('font:bold True;  borders:top double;', num_format_str='#,##0.00')

class PurchaseRequestReportOut(models.Model):
    _name = 'purchase.request.report.out'
    _description = 'purchase request report'

    purchase_request_data = fields.Char('Name', size=256)
    file_name = fields.Binary('PR Excel Report', readonly=True)
    purchase_request_work = fields.Char('Name', size=256)
    file_names = fields.Binary('PR CSV Report', readonly=True)


class WizardWizards(models.Model):
    _name = 'wizard.purchase.request.reports'
    _description = 'purchase request wizard'

    @api.multi
    def action_purchase_request_report(self):
        # XLS report
        rows = {}
        label_lists = ['PR NUMBER', 'ORIGIN', 'DATE', 'REQUESTED BY', 'APPROVAL', 'DESCRIPTION', 'COMPANY', 'PRODUCT',
                       'UOM',
                       'QTY', 'DESCR', 'SPEC']
        order = self.env['purchase.request'].browse(self._context.get('active_ids', list()))
        workbook = xlwt.Workbook()
        for rec in order:
            row = []
            for line in rec.line_ids:
                product = {'name': line.product_id.name, 'product_uom_id': line.product_uom_id.name,
                           'product_qty': line.product_qty, 'description': line.name,
                           'specifications': line.specifications, 'estimated_cost': line.estimated_cost,
                           'estimated_cost1': line.estimated_cost1, 'estimated_cost2': line.estimated_cost2
                           }
                row.append(product)

            rows['products'] = row
            rows['name'] = rec.name
            rows['origin'] = rec.origin
            rows['date_start'] = rec.date_start
            rows['requested_by'] = rec.requested_by.name
            rows['assigned_to'] = rec.assigned_to.name
            rows['description'] = rec.description
            rows['company_id'] = rec.company_id.name
            rows['net_amount_total'] = rec.net_amount_total
            rows['est_amount_total1'] = rec.est_amount_total1
            rows['est_amount_total2'] = rec.est_amount_total2


            sheet = workbook.add_sheet(rec.name, cell_overwrite_ok=True)

            sheet.write_merge(2, 3, 4, 9, 'Quotation Comparison Form', style2)
            sheet.write_merge(5, 5, 1, 2, 'No. :', style3)
            sheet.write_merge(7, 7, 1, 3, 'No BPPB & Date :', style3)
            sheet.write_merge(7, 7, 4, 5, rows['name'], style3)
            sheet.write_merge(8, 8, 1, 3, 'Head Office / Branch Office :', style3)
            sheet.write_merge(8, 8, 4, 5, rows['company_id'], style3)
            sheet.write_merge(5, 5, 8, 9, 'Date :', style3)
            sheet.write_merge(5, 5, 10, 11, rows['date_start'], style0)


            sheet.write_merge(11, 13, 1, 1, 'NO', style1)
            sheet.write_merge(11, 13, 2, 4, 'MATERIAL / ITEM', style1)
            sheet.write_merge(11, 13, 5, 5, 'QTY', style1)
            sheet.write_merge(11, 11, 6, 11, 'SUPPLIER NAME', style1)
            sheet.write_merge(12, 12, 6, 7, 'SUPPLIER1', style1)
            sheet.write_merge(12, 12, 8, 9, 'SUPPLIER2', style1)
            sheet.write_merge(12, 12, 10, 11, 'SUPPLIER3', style1)
            sheet.write(13, 6, 'STUAN', style1)
            sheet.write(13, 7, 'TOTAL', style1)
            sheet.write(13, 8, 'STUAN', style1)
            sheet.write(13, 9, 'TOTAL', style1)
            sheet.write(13, 10, 'STUAN', style1)
            sheet.write(13, 11, 'TOTAL', style1)
            sheet.write_merge(26, 26, 1, 5, 'GRANT TOTAL', style1)
            sheet.write_merge(26, 26, 6, 7, rows['net_amount_total'], style1)
            sheet.write_merge(26, 26, 8, 9, rows['est_amount_total1'], style1)
            sheet.write_merge(26, 26, 10, 11, rows['est_amount_total2'], style1)

            n = 14
            i = 1
            for product in rows['products']:
                sheet.write(n, 1, i, style5)
                sheet.write_merge(n, n, 2, 4, product['name'], style6)
                sheet.write(n, 5, product['product_qty'], style0)
                sheet.write(n, 6, product['product_uom_id'], style0)
                sheet.write(n, 7, product['estimated_cost'], style0)
                sheet.write(n, 8, product['product_uom_id'], style0)
                sheet.write(n, 9, product['estimated_cost1'], style0)
                sheet.write(n, 10, product['product_uom_id'], style0)
                sheet.write(n, 11, product['estimated_cost2'], style0)
                n += 1
                i += 1

        # CSV report
        datas = []
        for values in order:
            for value in values.line_ids:
                if value.product_id.seller_ids:
                    item = [
                        str(value.request_id.name or ''),
                        str(value.request_id.origin or ''),
                        str(value.request_id.date_start or ''),
                        str(value.request_id.requested_by.name or ''),
                        str(value.request_id.assigned_to.name or ''),
                        str(value.request_id.description or ''),
                        str(value.request_id.company_id.name or ''),
                        str(value.product_id.name or ''),
                        str(value.product_uom_id.name or ''),
                        str(value.product_qty or ''),
                        str(value.product_qty or ''),
                        str(value.description or ''),
                        str(value.specifications or ''),
                    ]
                    datas.append(item)

        output = StringIO()
        label = (';'.join(label_lists))
        output.write(label)
        output.write("\n")

        for data in datas:
            record = ';'.join(data)
            output.write(record)
            output.write("\n")
        data = base64.b64encode(bytes(output.getvalue(), "utf-8"))

        if platform.system() == 'Linux':
            filename = ('/tmp/PurchaseRequestReport-' + str(datetime.today().date()) + '.xls')
            filename2 = ('/tmp/PurchaseRequestReport-' + str(datetime.today().date()) + '.csv')
        else:
            filename = ('PurchaseRequestBPPB-' + str(datetime.today().date()) + '.xls')
            filename2 = ('PurchaseRequestReport-' + str(datetime.today().date()) + '.csv')

        filename = filename.split('/')[0]
        filename2 = filename2.split('/')[0]
        workbook.save(filename)
        fp = open(filename, "rb")
        file_data = fp.read()
        out = base64.encodestring(file_data)

        # Files actions
        attach_vals = {
            'purchase_request_data': filename,
            'file_name': out,
            'purchase_request_work': filename2,
            'file_names': data,
        }

        act_id = self.env['purchase.request.report.out'].create(attach_vals)
        fp.close()
        return {
            'type': 'ir.actions.act_window',
            'res_model': 'purchase.request.report.out',
            'res_id': act_id.id,
            'view_type': 'form',
            'view_mode': 'form',
            'context': self.env.context,
            'target': 'new',
        }
