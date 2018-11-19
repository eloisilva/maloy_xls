#! /usr/bin/env python3
#################################################################################
#     File Name           :     sigitm_readTP.py
#     Created By          :     Eloi Silva
#     Creation Date       :     [2017-07-30 23:24]
#     Last Modified       :     [2018-09-09 22:14]
#     Description         :      
#################################################################################

import xlrd, xlsxwriter, string, sys

def xls_to_list(file_xls, offset=0, sheet_index=0):
    # Open xls(x) file
    workbook = xlrd.open_workbook(file_xls)
    # Get sheet
    worksheet = workbook.sheet_by_index(sheet_index)
    rows = []
    for r in range(worksheet.nrows):
        if r < offset: continue
        cols = []
        for c in range(worksheet.ncols):
            cols.append(worksheet.cell_value(r, c))
        rows.append(cols)
    return rows

def sheet_names(file_xls):
    workbook = xlrd.open_workbook(file_xls)
    return workbook.sheet_names()

def list_to_xls(list_xls, file_xls, name=None, debug=0):
    # Add sheet name
    sheet_name = 'Sheet0' if not name else name
    # Create XLS(x) file
    workbook = xlsxwriter.Workbook(file_xls)
    worksheet = workbook.add_worksheet(sheet_name)
    # Add Header
    if debug > 0: print('\r[*]Adding Header to file: %s' % file_xls)
    format_header = workbook.add_format({'bold': True})
    format_header.set_align('center')
    format_header.set_align('vcenter')
    format_header.set_border(1)
    format_header.set_bg_color('#538DD5')
    header_range = 'A1:' + calc_integer_to_row_position(len(list_xls[0])-1) + '1'
    # Add Header
    for i in range(len(list_xls[0])):
        data = list_xls[0][i]
        field_size = len(data) + 5
        worksheet.set_column(0, i, field_size)
        worksheet.write(0, i, data, format_header)
        if debug > 0: print(data, end=',')
    worksheet.autofilter(header_range)
    if debug > 0: print('\n[+]Header added')
    # Add data
    for (ir, r) in enumerate(list_xls):
        if ir == 0: continue
        if debug > 0:
            sys.stdout.write('\r[*]Writeing row[%s]' % ir)
        for (ic, c) in enumerate(r):
            worksheet.write(ir, ic, c)
            if debug > 0:
                sys.stdout.flush()
                sys.stdout.write('\r[*]Writeing row[%s],col[%s]' % (ir, ic))
        if debug > 0:
            sys.stdout.flush()
            sys.stdout.write('\r[+]row[%s] | %s' % (ir, r))
            print()
    workbook.close()

def xls_rows(file_list):
    rows = []
    for (r, row) in enumerate(file_list):
        for (c, col) in enumerate(row):
            if col:
                rows.append(r)
                break
    return rows

def xls_cols(file_list, row):
    cols = []
    for (c, col) in enumerate(file_list[row]):
        if col:
            cols.append(c)
    return cols

def xls_clean(file, offset=0, sheet_index=0):
    file_list = xls_to_list(file, offset, sheet_index)
    rows = xls_rows(file_list)
    cols = xls_cols(file_list, rows[0])
    ans = []
    for (i, r) in enumerate(rows):
        ans.append([])
        for c in cols:
            ans[i].append(file_list[r][c])
    return ans

def calc_integer_to_row_position(x, base=26):
    if x > 1351:
        print('Error: Number to high. Try from 1 to 1351')
    elif x < base+1:
        return string.ascii_uppercase[x-1]
    else:
        bits = 1
        while x > base**(bits+1):
            bits += 1
        n = x // (base**bits)
        nl = int(x % (base**bits))
        if x / (base**bits) == n:
            return calc_integer_to_row_position(n-1)+calc_integer_to_row_position(nl)        
        return calc_integer_to_row_position(n)+calc_integer_to_row_position(nl)
