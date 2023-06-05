import gspread
import gspread_formatting
import time
import string


def isDebit(num):
    return not (200 <= num <= 499)


fmtDown = gspread_formatting.CellFormat(
    borders=gspread_formatting.Borders(
        # top=gspread_formatting.Border('SOLID_THICK'),
        bottom=gspread_formatting.Border('SOLID_THICK'),
        # left=gspread_formatting.Border('SOLID_THICK'),
        # right=gspread_formatting.Border('SOLID_THICK')
    )
)

fmtBar = gspread_formatting.CellFormat(
    borders=gspread_formatting.Borders(
        # top=gspread_formatting.Border('SOLID_THICK'),
        # bottom=gspread_formatting.Border('SOLID_THICK'),
        # left=gspread_formatting.Border('SOLID_THICK'),
        right=gspread_formatting.Border('SOLID_THICK')
    )
)

fmtSemi = gspread_formatting.CellFormat(
    borders=gspread_formatting.Borders(
        top=gspread_formatting.Border('SOLID'),
        bottom=gspread_formatting.Border('SOLID'),
        # left=gspread_formatting.Border('SOLID_THICK'),
        # right=gspread_formatting.Border('SOLID_THICK')
    )
)


def divmod_excel(n):
    a, b = divmod(n, 26)
    if b == 0:
        return a - 1, b + 26
    return a, b


def to_excel_char(c):
    chars = []
    while c > 0:
        c, d = divmod_excel(c)
        chars.append(string.ascii_uppercase[d - 1])
    return ''.join(reversed(chars))


def to_excel(r, c):
    chars = []
    while c > 0:
        c, d = divmod_excel(c)
        chars.append(string.ascii_uppercase[d - 1])
    return ''.join(reversed(chars)) + str(r)


def to_excel_range(a, b, c, d):
    return to_excel(a, b) + ':' + to_excel(c, d)


def writeToFile(file, num, a):
    file.write(f"Transaction {num} - {a}\n")


def boldRange(a, b, c, d):
    global wks
    wks.format(to_excel_range(a, b, c, d), {"textFormat": {"bold": True}})


sa = gspread.service_account()
sh = sa.open("Aaron - #3 Monopoly T-Accounts")
wks = sh.worksheet("T-Accounts")
# wks.clear()

# gspread_formatting.format_cell_range(wks, to_excel(6, 6), fmt)


values = []
with open("formatted-journal.txt") as f:
    tmp = []
    for line in [x.rstrip() for x in f.readlines()]:
        if line == '':
            values.append(tmp)
            tmp = []
        else:
            tmp.append(line)

with open("accounts") as f:
    accounts = [x.rstrip() for x in f.readlines()]

startRow = 2
startCol = 2

startAt = ""
start = 0

for q in range(0, len(accounts), 2):
    account_num = int(accounts[q])
    account_name = accounts[q + 1]
    if account_name == "Income Summary":
        continue
    print(account_name, account_num)
    if startAt == account_name or startAt == "":
        start = 1
    if not start:
        startCol += 4
        continue

    debitFile = open(f"t-accounts/{account_name}-debit.txt", "w")
    creditFile = open(f"t-accounts/{account_name}-credit.txt", "w")

    wks.merge_cells(startRow, startCol, startRow, startCol + 1)
    wks.format(to_excel(startRow, startCol), {"horizontalAlignment": "CENTER", "textFormat": {"bold": True}})
    wks.format(to_excel_range(startRow + 1, startCol - 1, 999, startCol), {"horizontalAlignment": "RIGHT"})
    wks.format(to_excel_range(startRow + 1, startCol + 1, 999, startCol + 2), {"horizontalAlignment": "LEFT"})
    boldRange(startRow, startCol - 1, 999, startCol - 1)
    boldRange(startRow, startCol + 2, 999, startCol + 2)
    gspread_formatting.set_column_width(wks, to_excel_char(startCol - 1), 40)
    gspread_formatting.set_column_width(wks, to_excel_char(startCol + 2), 40)

    # add borders
    wks.update_cell(startRow, startCol, account_name)

    debitRow = startRow + 1
    creditRow = startRow + 1

    time.sleep(10)

    for i in range(len(values) - 3):  # exclude closing entries
        transaction = values[i]
        for j in range(len(transaction)):
            if str(transaction[j]).strip() == account_name:
                if transaction[j].strip() != transaction[j]:  # if credit
                    cells = wks.range(to_excel_range(creditRow, startCol + 1, creditRow, startCol + 2))
                    cells[0].value = int(transaction[j + 1])
                    cells[1].value = i + 1 if i > 1 else "Bal."
                    # wks.update_cell(creditRow, startCol + 1, int(transaction[j + 1]))  # transaction
                    # if i > 1: wks.update_cell(creditRow, startCol + 2, i + 1)  # transaction number
                    # else: wks.update_cell(creditRow, startCol + 2, "Bal.")  # transaction number

                    creditRow += 1
                    writeToFile(creditFile, i + 1, transaction[j + 1])
                    wks.update_cells(cells, value_input_option='USER_ENTERED')
                else:
                    cells = wks.range(to_excel_range(debitRow, startCol - 1, debitRow, startCol))
                    cells[0].value = i + 1 if i > 1 else "Bal."
                    cells[1].value = int(transaction[j + 1])
                    # wks.update_cell(debitRow, startCol, int(transaction[j + 1]))
                    # if i > 1: wks.update_cell(debitRow, startCol - 1, i + 1)  # transaction number
                    # else: wks.update_cell(debitRow, startCol - 1, "Bal.")  # transaction number
                    debitRow += 1
                    writeToFile(debitFile, i + 1, transaction[j + 1])
                    wks.update_cells(cells, value_input_option='USER_ENTERED')
                time.sleep(1)

    # underline formatting
    gspread_formatting.format_cell_range(wks, to_excel_range(startRow, startCol, startRow, startCol + 1), fmtDown)
    gspread_formatting.format_cell_range(wks, to_excel_range(startRow + 1, startCol, 5 + max(debitRow, creditRow), startCol), fmtBar)

    next = max(debitRow, creditRow)
    gspread_formatting.format_cell_range(wks, to_excel_range(next, startCol, next, startCol + 1), fmtSemi)

    cells = wks.range(to_excel_range(next, startCol, next + 1, startCol + 1))
    cells[0].value = f"=SUM({to_excel_range(startRow + 1, startCol, next - 1, startCol)})"
    cells[1].value = f"=SUM({to_excel_range(startRow + 1, startCol + 1, next - 1, startCol + 1)})"
    if isDebit(account_num):
        cells[2].value = f"={to_excel(next, startCol)} - {to_excel(next, startCol + 1)}"
    else:
        cells[3].value = f"={to_excel(next, startCol + 1)} - {to_excel(next, startCol)}"
    wks.update_cells(cells, value_input_option='USER_ENTERED')

    # wks.update_cell(next, startCol, f"=SUM({to_excel_range(startRow + 1, startCol, next - 1, startCol)})")
    # wks.update_cell(next, startCol + 1, f"=SUM({to_excel_range(startRow + 1, startCol + 1, next - 1, startCol + 1)})")
    #
    # if isDebit(account_num):
    #     wks.update_cell(next + 1, startCol, f"={to_excel(next, startCol)} - {to_excel(next, startCol + 1)}")
    # else:
    #     wks.update_cell(next + 1, startCol + 1, f"={to_excel(next, startCol + 1)} - {to_excel(next, startCol)}")
    boldRange(next + 1, startCol, next + 1, startCol + 1)


    startCol += 4
    debitFile.close()
    creditFile.close()

    time.sleep(2)
