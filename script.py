from __future__ import print_function
import os
import xlrd
from mailmerge import MailMerge
from datetime import date

template = 'templates/invoice_template.docx'
excel_sheet = xlrd.open_workbook('data_source/invoice_data.xlsx').sheet_by_index(0)

document = MailMerge(template)

def get_invoice_info(excel_sheet):
    invoice_info = {
        'Name': excel_sheet.cell(1,1).value,
        'StreetAddress': excel_sheet.cell(2,1).value,
        'CityStateZip': excel_sheet.cell(3,1).value,
        'InvoiceDate': '{:%b %d %Y}'.format(date.today()),
        'PayerName': excel_sheet.cell(6,1).value,
        'PayerStreetAddress': excel_sheet.cell(7,1).value,
        'PayerCityStateZip': excel_sheet.cell(8,1).value,
        'TotalBilled': '$4000.00'
    }
    return invoice_info

def get_work_events(excel_sheet):
    work_events = []
    total_billed = 0
    for i in range(11,excel_sheet.nrows):
        row_values = excel_sheet.row_values(i)
        year,month,day = xlrd.xldate_as_tuple(row_values[0],0)[:3]
        title = row_values[1]
        duration = row_values[2]
        total_billed += duration*150
        description = row_values[3]
        work_event = {
            'Date': '{}/{}/{}'.format(month,day,year),
            'WorkTitle': title,
            'Duration': str(duration),
            'WorkDescription': description,
            'AmountBilled': '${}.00'.format(str(int(duration)*150))
        }
        work_events.append(work_event)
    return work_events, total_billed

invoice_info = get_invoice_info(excel_sheet)
work_events, total_billed = get_work_events(excel_sheet)
invoice_info['TotalBilled'] = '${}.00'.format(str(int(total_billed)))

document.merge(**invoice_info)
document.merge_rows('WorkTitle',work_events)
document.write('reports/Invoice - {}.docx'.format(date.today()))
