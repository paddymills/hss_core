import xlwings

wb = xlwings.books["rawcogi.xlsx"]

cohv = list()
for c in wb.sheets[2].range("A2").expand().value:
    cohv.append([c[0], c[1]])  # Order, PartName

mb51 = list()
for m in wb.sheets[2].range("A2").expand().value:
    mb51.append([m[0], m[4], m[5]])  # MM, Order, Quantity

migo = list()
for c in wb.sheets[0].range("A2").expand().value:
    cogiMM = c[0]
    cogiOrder = c[9]
    for x, y in cohv:
        if x == cogiOrder:
            part = y
            break

    for x, y in cohv:
        if y == part and x != cogiOrder:
            otherOrder = x
            break

    for x, y, z in mb51:
        if y == otherOrder and x == cogiMM and z >= c[15]:
            migo.append([x, None, c[2], c[3], c[15]])
            break
    else:
        print(c[0], "not matched")

wb.sheets[4].range("A2").value = migo
