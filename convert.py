#!/usr/bin/env python3

import csv
import datetime
import json
import openpyxl


# header font
HFONT_STYLE = openpyxl.styles.Font(bold=True)

# border line styles
SIDE_BLACK_THIN = openpyxl.styles.borders.Side(style='thin', color='000000')
SIDE_BLACK_DOUBLE = openpyxl.styles.borders.Side(style='double', color='000000')

# cell border styles
HBORDER_STYLE = openpyxl.styles.borders.Border(top=SIDE_BLACK_THIN, bottom=SIDE_BLACK_DOUBLE, left=SIDE_BLACK_THIN, right=SIDE_BLACK_THIN)
BORDER_STYLE = openpyxl.styles.borders.Border(bottom=SIDE_BLACK_THIN, left=SIDE_BLACK_THIN, right=SIDE_BLACK_THIN)
MBORDER_STYLE = openpyxl.styles.borders.Border(top=SIDE_BLACK_THIN, bottom=SIDE_BLACK_THIN, left=SIDE_BLACK_THIN, right=SIDE_BLACK_THIN)

# fill styles: certificate type
FILL_CERT_TYPE_ROOT = openpyxl.styles.PatternFill(patternType='solid', fgColor='fcf3d0', bgColor='fcf3d0')
FILL_CERT_TYPE_INTERMEDIATE = openpyxl.styles.PatternFill(patternType='solid', fgColor='dceaf6', bgColor='dceaf6')

# fill style: revoked certificates
FILL_REVOKED = openpyxl.styles.PatternFill(patternType='solid', fgColor='ff3333', bgColor='ff3333')

# fill style: technically-constrained certificates
FILL_TECHNICALLY_CONSTRAINED = openpyxl.styles.PatternFill(patternType='solid', fgColor='e9f3ec', bgColor='e9f3ec')


def canonicalize(row):
    # Audits Same as Parent?
    row[9] = row[9].upper()
    # CP/CPS Same as Parent?
    row[31] = row[31].upper()
    # Technically Constrained
    row[38] = row[38].upper()
    # date (YYYY-mm-dd)
    row[13] = row[13].replace('.', '-')
    row[14] = row[14].replace('.', '-')
    row[15] = row[15].replace('.', '-')
    row[18] = row[18].replace('.', '-')
    row[19] = row[19].replace('.', '-')
    row[20] = row[20].replace('.', '-')
    row[23] = row[23].replace('.', '-')
    row[24] = row[24].replace('.', '-')
    row[25] = row[25].replace('.', '-')
    row[28] = row[28].replace('.', '-')
    row[29] = row[29].replace('.', '-')
    row[30] = row[30].replace('.', '-')
    row[34] = row[34].replace('.', '-')
    # JSON array
    if row[43] != '':
        row[43] = '\n'.join(json.loads(row[43]))


def add_metadata_sheet(metadata_sheet):
    metadata_sheet.column_dimensions['A'].width = 16
    metadata_sheet.column_dimensions['B'].width = 48

    row = [
        'source',
        'https://www.ccadb.org/resources',
    ]
    row = [openpyxl.cell.WriteOnlyCell(metadata_sheet, value=c) for c in row]
    for cell in row:
        cell.number_format = openpyxl.styles.numbers.FORMAT_TEXT
        cell.border = MBORDER_STYLE
    metadata_sheet.append(row)

    row = [
        'generated at',
        datetime.datetime.now(tz=datetime.timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
    ]
    row = [openpyxl.cell.WriteOnlyCell(metadata_sheet, value=c) for c in row]
    for cell in row:
        cell.number_format = openpyxl.styles.numbers.FORMAT_TEXT
        cell.border = MBORDER_STYLE
    metadata_sheet.append(row)


def main():
    num_records = 0
    with open('AllCertificateRecordsCSVFormat', 'r', encoding='UTF-8', newline='') as csv_fh:
        csv_reader = csv.reader(csv_fh, dialect='unix')
        for _ in csv_reader:
            num_records += 1

    book = openpyxl.Workbook(write_only=True)

    sheet = book.create_sheet(title='AllCertificateRecords')

    sheet.auto_filter.ref = f"A1:AT{num_records}"
    sheet.freeze_panes = 'D2'
    sheet.column_dimensions['A'].width = 14
    sheet.column_dimensions['B'].width = 4
    sheet.column_dimensions['C'].width = 36
    sheet.column_dimensions['D'].width = 4
    sheet.column_dimensions['E'].width = 24
    sheet.column_dimensions['F'].width = 16
    sheet.column_dimensions['G'].width = 16
    sheet.column_dimensions['H'].width = 4
    sheet.column_dimensions['I'].width = 4
    sheet.column_dimensions['J'].width = 8
    sheet.column_dimensions['K'].width = 24
    sheet.column_dimensions['L'].width = 14
    sheet.column_dimensions['M'].width = 14
    sheet.column_dimensions['N'].width = 12
    sheet.column_dimensions['O'].width = 12
    sheet.column_dimensions['P'].width = 12
    sheet.column_dimensions['Q'].width = 14
    sheet.column_dimensions['R'].width = 14
    sheet.column_dimensions['S'].width = 12
    sheet.column_dimensions['T'].width = 12
    sheet.column_dimensions['U'].width = 12
    sheet.column_dimensions['V'].width = 14
    sheet.column_dimensions['W'].width = 14
    sheet.column_dimensions['X'].width = 12
    sheet.column_dimensions['Y'].width = 12
    sheet.column_dimensions['Z'].width = 12
    sheet.column_dimensions['AA'].width = 14
    sheet.column_dimensions['AB'].width = 14
    sheet.column_dimensions['AC'].width = 12
    sheet.column_dimensions['AD'].width = 12
    sheet.column_dimensions['AE'].width = 12
    sheet.column_dimensions['AF'].width = 8
    sheet.column_dimensions['AG'].width = 14
    sheet.column_dimensions['AH'].width = 14
    sheet.column_dimensions['AI'].width = 12
    sheet.column_dimensions['AJ'].width = 14
    sheet.column_dimensions['AK'].width = 14
    sheet.column_dimensions['AL'].width = 14
    sheet.column_dimensions['AM'].width = 8
    sheet.column_dimensions['AN'].width = 18
    sheet.column_dimensions['AO'].width = 18
    sheet.column_dimensions['AP'].width = 24
    sheet.column_dimensions['AQ'].width = 14
    sheet.column_dimensions['AR'].width = 14
    sheet.column_dimensions['AS'].width = 8
    sheet.column_dimensions['AT'].width = 4

    cert_type_rules = [
        openpyxl.formatting.rule.CellIsRule(
            operator='equal',
            formula=['"Root Certificate"'],
            stopIfTrue=False,
            fill=FILL_CERT_TYPE_ROOT
            ),
        openpyxl.formatting.rule.CellIsRule(
            operator='equal',
            formula=['"Intermediate Certificate"'],
            stopIfTrue=False,
            fill=FILL_CERT_TYPE_INTERMEDIATE
            ),
    ]
    for rule in cert_type_rules:
        sheet.conditional_formatting.add(f"F2:F{num_records}", rule)

    cert_revoked_rules = [
        openpyxl.formatting.rule.CellIsRule(
            operator='equal',
            formula=['"Revoked"'],
            stopIfTrue=False,
            fill=FILL_REVOKED
            ),
        openpyxl.formatting.rule.CellIsRule(
            operator='equal',
            formula=['"Parent Cert Revoked"'],
            stopIfTrue=False,
            fill=FILL_REVOKED
            ),
    ]
    for rule in cert_revoked_rules:
        sheet.conditional_formatting.add(f"G2:G{num_records}", rule)

    cert_constrained_rules = [
        openpyxl.formatting.rule.CellIsRule(
            operator='equal',
            formula=['"TRUE"'],
            stopIfTrue=False,
            fill=FILL_TECHNICALLY_CONSTRAINED
            ),
    ]
    for rule in cert_constrained_rules:
        sheet.conditional_formatting.add(f"AM2:AM{num_records}", rule)

    with open('AllCertificateRecordsCSVFormat', 'r', encoding='UTF-8', newline='') as csv_fh:
        csv_reader = csv.reader(csv_fh, dialect='unix')
        header = next(csv_reader)
        header.append('X-Number of items in "JSON Array of Partitioned CRLs"')
        header.append('X-crt.sh link')
        header = [openpyxl.cell.WriteOnlyCell(sheet, value=hc) for hc in header]
        for hc in header:
            hc.font = HFONT_STYLE
            hc.border = HBORDER_STYLE
            hc.number_format = openpyxl.styles.numbers.FORMAT_TEXT
        sheet.row_dimensions[1].height = 14.25
        sheet.append(header)

        for row_no, row in enumerate(csv_reader, 2):
            if len(row) != 44:
                raise RuntimeError(f"unexpected number of rows {len(row)} at CSV line {row_no}")
            canonicalize(row)

            # X-Number of items in "JSON Array of Partitioned CRLs"
            if row[43] != '':
                row.append(row[43].count('\n') + 1)
            else:
                row.append('')

            # X-crt.sh link
            row.append(f"https://crt.sh/?sha256={row[7]}")

            row = [openpyxl.cell.WriteOnlyCell(sheet, value=c) for c in row]
            for col_idx, cell in enumerate(row):
                cell.border = BORDER_STYLE
                if col_idx in {9, 31, 38}:
                    cell.number_format = openpyxl.styles.numbers.FORMAT_GENERAL
                elif col_idx in {13, 14, 15, 18, 19, 20, 23, 24, 25, 28, 29, 30, 34}:
                    cell.number_format = openpyxl.styles.numbers.FORMAT_DATE_YYYYMMDD2
                    if cell.value != '':
                        cell.value = datetime.date.fromisoformat(cell.value)
                    else:
                        cell.value = None
                elif col_idx in {44}:
                    cell.number_format = openpyxl.styles.numbers.FORMAT_NUMBER
                elif col_idx in {45}:
                    cell.number_format = openpyxl.styles.numbers.FORMAT_TEXT
                    href = cell.value
                    cell.value = '\U0001F4DC'
                    cell.hyperlink = href
                else:
                    cell.number_format = openpyxl.styles.numbers.FORMAT_TEXT
            sheet.row_dimensions[row_no].height = 13.5
            sheet.append(row)

    add_metadata_sheet(book.create_sheet(title='_metadata'))

    book.active = book.worksheets[0]

    book.save('CCADB-certificates.xlsx')


if __name__ == '__main__':
    main()


# vim: set fileencoding=utf-8 nobomb fileformat=unix filetype=python number expandtab tabstop=8 softtabstop=4 shiftwidth=4 autoindent smartindent :
