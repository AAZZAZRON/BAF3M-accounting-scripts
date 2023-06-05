accounts = {}
with open('accounts') as f:
    account_lines = [x.strip() for x in f.readlines()]
    for i in range(0, len(account_lines), 2):
        accounts[account_lines[i + 1]] = int(account_lines[i])


with open("raw-journal-data.txt") as f:
    lines = [x.rstrip() for x in f.readlines() if x != '\n']


values = []
arr = []
tmp = ""
pageNum = 0
for line in lines:
    if line.find("Page No: ") == 0:
        pageNum = int(line[-2:])
        continue
    if line in ["Date", "Particulars", "PR", "DR", "CR"]:
        continue
    elif line.isdigit() and len(line) != 3:
        tmp += line
    elif line.isdigit() and len(line) == 3:
        assert accounts[arr[-1].strip()] == int(line)
        continue
    elif tmp != "" and line == "-":
        arr.append(int(tmp))
        tmp = ""
    else:
        if tmp != "":
            values.append(arr)
            arr = [tmp[-2:], pageNum]
            tmp = ""
        arr.append(line)

values.pop(0)
values.append(arr)

out = open('formatted-journal.txt', 'w')
for transaction in values:
    for x in transaction:
        if "â€™" in str(x):
            x = str(x).replace("â€™", "'")
        out.write(f"{x}\n")
    out.write("\n")
out.close()
