from openpyxl import load_workbook
wb = load_workbook("TEMPLATE HARGA MAS2.xlsx")
ws = wb["EMAS"]
print("A1:", repr(ws["A1"].value), "E1:", repr(ws["E1"].value))
print("--- kolom A (label) + B/D/F/H ---")
for r in range(1, 30):
    a = ws.cell(r, 1).value
    b = ws.cell(r, 2).value
    d = ws.cell(r, 4).value
    f = ws.cell(r, 6).value
    h = ws.cell(r, 8).value
    print(r, repr(a), "| B:", repr(b), "| D:", repr(d), "| F:", repr(f), "| H:", repr(h))
