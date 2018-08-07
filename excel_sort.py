import xlsxwriter
import xlrd
from datetime import datetime

# style0 = xlwt.easyxf('font : name Time New Roman, color-index red, bold on', num_format_str='#,##0.00')

# style1 = xlwt.easyxf(num_format_str='D-MMM-YY')

# This Creates the new Workbook
wb = xlsxwriter.Workbook('Book1.xlsx')
ws = wb.add_worksheet("Sheet 1")

# This opens the Catalog
xl_workbook = xlrd.open_workbook("cat.xlsx")
xl_sheet = xl_workbook.sheet_by_index(0)

c = xl_sheet.col(0)
r = xl_sheet.row(0)
# col = xl_sheet.ncols
# int(col)
# col -= 1
# i = 0
# while i < col:
#     val = xl_sheet.col(0)[i].value
#     ws.write(0, i, 'val')/
#     i += 1

# This writes first row of spreadsheet

for x in range(len(r)):
    dat = xl_sheet.cell(0, x)
    ws.write(0, x, dat.value)


def get_col():
    row = input("Enter Column To Search (e.g. 'AB'): ")
    row = row.lower()
    s_row = list(row)
    if len(s_row) == 1:
        return ord(s_row[0]) - 97
    else:
        first = ord(s_row[0]) - 96
        second = ord(s_row[1]) - 97
        num = first * 26
        num += second
        return num


def get_term():
    term = input("Enter Full Search Term: ")
    # term = term.lower()
    q = input("Is the search term a number? (y/n): ")
    if q.lower() == 'y':
        term = float(term)
    return term


def search_doc(row, col, s_term, num_col, wss):
    # index for new Book
    w = 1

    for i in range(len(row)):
        val = xl_sheet.cell(i, col)
        string1 = val.value
        if string1 == s_term:
            for j in range(len(num_col)):
                data = xl_sheet.cell(i, j)
                # print(j)
                wss.write(w, j, data.value)
            w = w + 1
    print(w - 1, "Entries found in file")


search_c = get_col()
word = get_term()

search_doc(c, search_c, word, r, ws)

wb.close()
