accounts = {}
with open('accounts') as f:
    lines = [x.strip() for x in f.readlines()]
    for i in range(0, len(lines), 2):
        accounts[lines[i + 1]] = int(lines[i])

print(accounts)

values = []
with open("formatted-journal.txt") as f:
    tmp = []
    for line in [x.rstrip() for x in f.readlines()]:
        if line == '':
            values.append(tmp)
            tmp = []
        else:
            tmp.append(line)

page = 0
out = None
for transaction in values:
    if transaction[1] != page:
        page = transaction[1]
        out = open(f'journal-PRs/{page}.txt', 'w')
    out = open(f'journal-PRs/{page}.txt', 'a')
    for line in transaction:
        if line.strip() in accounts:
            print(accounts[line.strip()])
            out.write(str(accounts[line.strip()]) + '\n')
    out.write('\n\n')
    out.close()
