from __future__ import print_function
import os
from mailmerge import MailMerge
from datetime import date

template = 'templates/invoice_template.docx'

document = MailMerge(template)

invoice_info = {
    'Name': 'John Smith',
    'StreetAddress': '5555 FakeAddress Drive',
    'CityStateZip':'Austin, TX 78749',
    'InvoiceDate': '{:%b %d %Y}'.format(date.today()),
    'PayeeName': 'Kelly Smith',
    'PayeeStreetAddress': '5555 FakeAddress Drive',
    'PayeeCityStateZip': 'Austin, TX 78749',
    'TotalBilled': '$4000.00'
}

document.merge(**invoice_info)
document.write('reports/testoutput.docx')
