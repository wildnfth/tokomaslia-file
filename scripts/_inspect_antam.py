import openpyxl
wb = openpyxl.load_workbook(r"D:\TOKO MAS LIA\DOKUMEN\TEMPLATE HARGA MAS2.xlsx", data_only=False)
ws = wb["ANTAM"]
print("dims:", ws.dimensions, "max_row:", ws.max_row, "max_col:", ws.max_column)
# print last ~40 rows, columns A..L
for r in range(1, ws.max_row+1):
    vals = []
    for c in range(1, 13):
        v = ws.cell(row=r, column=c).value
        if v is not None:
            vals.append(f"{openpyxl.utils.get_column_letter(c)}{r}={v!r}")
    if vals:
        print(" | ".join(vals))
