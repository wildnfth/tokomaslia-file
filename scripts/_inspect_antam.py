# -*- coding: utf-8 -*-
from openpyxl import load_workbook

path = r"D:\TOKO MAS LIA\DOKUMEN\TEMPLATE HARGA MAS2.xlsx"
wb = load_workbook(path, data_only=True)
ws = wb["ANTAM"]
headers = []
for r in range(1, ws.max_row + 1):
    v = ws.cell(r, 1).value
    if v and str(v).strip().upper() == "HARGA ANTAM":
        headers.append(r)
print("header_count", len(headers))
print("last_headers", headers[-3:] if len(headers) >= 3 else headers)
if headers:
    h = headers[-1]
    print("last_header_row", h)
    print("A", h, ws.cell(h, 1).value)
    print("A", h + 1, ws.cell(h + 1, 1).value)
    print("G", h, ws.cell(h, 7).value)
    print("G", h + 1, ws.cell(h + 1, 7).value)
    print("K", h, ws.cell(h, 11).value)
    print("K", h + 1, ws.cell(h + 1, 11).value)
    print("--- sample prices last block ---")
    for r in range(h, min(h + 20, ws.max_row + 1)):
        a = ws.cell(r, 1).value
        b = ws.cell(r, 2).value
        g = ws.cell(r, 7).value
        hval = ws.cell(r, 8).value
        k = ws.cell(r, 11).value
        l = ws.cell(r, 12).value
        if any(x not in (None, "") for x in (a, b, g, hval, k, l)):
            print(r, "A=", a, "B=", b, "G=", g, "H=", hval, "K=", k, "L=", l)
wb.close()
