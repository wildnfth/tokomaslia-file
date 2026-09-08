# Konversi tabel harga:
# Kuning:
# 6K: 808.500 / 831.600 -> 810 / 830
# 8K: 1.016.400 / 1.039.500 -> 1015 / 1040
# 9K: 1.155.000 / 1.178.100 -> 1155 / 1180
# 10K: 1.247.400 -> 1250
# 16K: 1.755.600 / 1.790.250 -> 1755 / 1790
# 17K: 1.917.300 / 1.951.950 -> 1920 / 1950 (atau 1915 / 1950? 1917.3 -> 1915 atau 1920? round(1917.3/5)*5 = round(383.46)*5 = 383*5 = 1915!)
# Mari kita buat script pembulatan standar round(x/5)*5

items_kuning = [
    ("300/6K", 808.5, 831.6),
    ("375/8K", 1016.4, 1039.5),
    ("420/9K", 1155.0, 1178.1),
    ("450/10K", 1247.4, None),
    ("700/16K", 1755.6, 1790.25),
    ("750/17K", 1917.3, 1951.95),
    ("18K", 2032.8, None),
    ("875/20K", 2159.85, None),
    ("9A/916/21K", 2217.6, None),
    ("9B", 2310.0, None),
    ("9C", 2425.5, None),
    ("9C PABRIK", 2564.1, None),
    ("9C KOPONG", 2610.3, None),
]

items_putih = [
    ("300/6K", 831.6),
    ("375/8K", 1062.6),
    ("420/9K", 1201.2),
    ("450/10K", 1293.6),
    ("700/16K", 1778.7),
    ("750/17K", 1963.5),
    ("18K", 2079.0),
]

def snap5(v):
    return int(round(v / 5.0) * 5)

print("KUNING:")
for label, v1, v2 in items_kuning:
    if v2 is not None:
        s1 = snap5(v1)
        s2 = snap5(v2)
        print(f'"{label}": "{s1} / {s2}",  # raw: {v1} / {v2}')
    else:
        s1 = snap5(v1)
        print(f'"{label}": {s1},  # raw: {v1}')

print("\nPUTIH:")
for label, v in items_putih:
    s = snap5(v)
    print(f'"{label}": {s},  # raw: {v}')
