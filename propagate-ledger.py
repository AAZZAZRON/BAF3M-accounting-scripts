import gspread
import time
import string

sa = gspread.service_account()
sh = sa.open("Aaron - #5 Monopoly Ledger")


def isDebit(num):
    return not (200 <= num <= 499)



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


values = []
with open("formatted-journal.txt") as f:
    tmp = []
    for line in [x.rstrip() for x in f.readlines()]:
        if line == '':
            values.append(tmp)
            tmp = []
        else:
            tmp.append(line)

with open("ledger-admin-data") as f:
    lines = [x.rstrip() for x in f.readlines()]

startAt = ""
start = 0

for q in range(6, len(lines), 6):
    account_name = lines[q]
    account_num = int(lines[q + 1])
    worksheet_name = lines[q + 2]
    row = int(lines[q + 3])
    col = int(lines[q + 4])
    print(account_name, worksheet_name, row, col)
    if startAt == account_name or startAt == "":
        start = 1
    if not start:
        continue
    wks = sh.worksheet(worksheet_name)

    file = open(f"ledger/{account_name}.txt", "w")

    cells = wks.range(to_excel_range(row - 2, col, row, col + 7))
    cells[0].value = "Account:"
    cells[1].value = account_name
    cells[6].value = "No:"
    cells[7].value = account_num
    cells[9].value = "Date"
    cells[10].value = "Particulars"
    cells[11].value = "PR"
    cells[12].value = "Debit"
    cells[13].value = "Credit"
    cells[14].value = "Dr/Cr"
    cells[15].value = "Balance"
    cells[16].value = "May"
    cells[-1].value = f"=E{row}" if isDebit(account_num) else f"=F{row}"
    wks.update_cells(cells, value_input_option='USER_ENTERED')
    # wks.update_cell(row - 2, col + 1, account_name)  # housekeeping stuff
    # wks.update_cell(row - 2, col + 7, account_num)
    # wks.update_cell(row, col, "May")
    # if isDebit(account_num):
    #     wks.update_cell(row, col + 7, f'=E{row}')
    # else:
    #     wks.update_cell(row, col + 7, f'=F{row}')
    time.sleep(1)
    for i in range(len(values)):
        transaction = values[i]
        for j in range(len(transaction)):
            if str(transaction[j]).strip() == account_name:
                cells = wks.range(to_excel_range(row, col + 1, row, col + 6))
                cells[0].value = transaction[0]
                cells[1].value = transaction[-1]
                cells[2].value = "J" + str(transaction[1])
                if transaction[j].strip() == transaction[j]:
                    cells[3].value = int(transaction[j + 1])
                    cells[4].value = 0
                else:
                    cells[3].value = 0
                    cells[4].value = int(transaction[j + 1])
                wks.update_cells(cells, value_input_option='USER_ENTERED')

                time.sleep(1)

                # wks.update_cell(row, col + 1, transaction[0])  # date
                # wks.update_cell(row, col + 2, transaction[-1])  # PR
                # wks.update_cell(row, col + 3, transaction[1])  # journal page number
                # if transaction[j].strip() == transaction[j]:
                #     wks.update_cell(row, col + 4, int(transaction[j + 1]))
                # else:
                #     wks.update_cell(row, col + 5, int(transaction[j + 1]))

                balance = wks.cell(row, col + 7).value.strip().replace(',', '')
                if balance != "-" and balance != "#REF!" and float(balance) > 0:
                    wks.update_cell(row, col + 6, "Dr" if isDebit(account_num) else "Cr")
                row += 1

                print(f"{i + 1}; {transaction[-1]}; {transaction[j]}; {transaction[j + 1]}")
                file.write(f"May {transaction[0]}\n")
                file.write(f"Transaction {i + 1} on journal page {transaction[1]}\n")
                file.write(f"PR: {transaction[-1]}\n")
                file.write(f"{transaction[j]} - {transaction[j + 1]}\n\n")
                time.sleep(1)
                # time.sleep(5)
    file.close()
