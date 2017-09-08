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

WorkEvent = {
        'Title': 'Texas Pathways Project, Institute #2',
        'Duration': 3,
        'Description': 'Assisted in lots of stuff.'
    }


work_history = [
    {
        'WorkDescription':WorkEvent['Title'],
        'Duration':str(WorkEvent['Duration']),
        'WorkDetails':WorkEvent['Description'],
        'AmountBilled': '${}.00'.format(str(WorkEvent['Duration']*150))
    }
]

document.merge(**invoice_info)
document.merge_rows('WorkDescription',work_history)
document.write('reports/Invoice - {}.docx'.format(date.today()))
