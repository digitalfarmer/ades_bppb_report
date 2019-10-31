# -*- coding: utf-8 -*-
# Part of Odoo. See LICENSE file for full copyright and licensing details.
{
    'name' : 'Purchase Request BPPB Excel Report',
    'version': '11.0',
    'author': 'adesdev',
    'category': 'Purchase',
    'license': 'LGPL-3',
    'support': 'ades@binasanprima.com',
    'website': 'https://digitalfarmer.github.io',
    'summary': 'Excel sheet for Purchase Request',
    'description': """ Purchase Request excel report
When user need to print the excel report in purchase request select the purchase request list and
user need to click the "Purchase request Excel Report" button and message will appear.select the "Print Excel report"button
for generating the purchase request excel file""",
    'depends': [
        'purchase_request','base'
    ],
    'data': [
        'security/ir.model.access.csv',
        'wizard/purchase_request_xls_view.xml',
    ],
    'installable': True,
    'application': True,
    'auto_install': False,
}
