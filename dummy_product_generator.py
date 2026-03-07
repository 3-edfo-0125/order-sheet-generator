from pathlib import Path
import csv
import random
import string

BASE_DIR = Path(__file__).resolve().parent

PRODUCT_QTY = 32
QTY_MIN = 15
QTY_MAX = 22

lst2 = []
for i in range(PRODUCT_QTY):
    al = random.choice(string.ascii_uppercase)
    n1 = random.randint(0, 9999)
    n2 = random.randint(0, 9999999)
    qty = random.randint(QTY_MIN, QTY_MAX) # 在庫数設定
    lst2.append([al + str(n1).zfill(4) + "-" + str(n2).zfill(7), qty])

lst2.sort()

with open(BASE_DIR / "product.csv", mode="w", encoding="utf_8", newline="") as f:
    writer = csv.writer(f)
    for i in range(len(lst2)):
        writer.writerow(lst2[i])
