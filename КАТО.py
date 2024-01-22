import xlrd
import csv

wb = xlrd.open_workbook(r'C:\Users\Danik\Downloads\КАТО_17.11.2023.xls')
ws = wb.sheet_by_index(0)

col = csv.writer(open("КАТО_old.csv", 'w', newline="", encoding="utf-8"))

for row in range(ws.nrows):
    col.writerow(ws.row_values(row))

with open("КАТО_old.csv", 'r', newline="", encoding="utf-8") as csv_file:
    reader = csv.DictReader(csv_file)
    rows = list(reader)


    output_fields = reader.fieldnames
    output_fields.remove("nn")
    output_fields.remove("ab")
    output_fields.remove("cd")
    output_fields.remove("ef")
    output_fields.remove("hij")
    output_fields.remove("k")

    output_fields.append("id")
    output_fields.append("parent_id")

    for i, row in enumerate(rows, start=1):
        row['id'] = i

        if float(row["ab"]) > 0 and float(row["cd"]) > 0 and float(row["ef"]) == 0 and float(row["hij"]) == 0:
            previous_i1 = i - 1
            row["parent_id"] = previous_i1
        elif float(row["ab"]) > 0 and float(row["cd"]) > 0 and float(row["ef"]) > 0 and float(row["hij"]) == 0:
            previous_i2 = i
            row["parent_id"] = previous_i1 + 1
        elif float(row["ab"]) > 0 and float(row["cd"]) > 0 and float(row["ef"]) > 0 and float(row["hij"]) > 0:
            row["parent_id"] = previous_i2
        else:
            row["parent_id"] = None

with open('КАТО_new.csv', 'w', newline="", encoding="utf-8") as csv_new:
    writer = csv.DictWriter(
        csv_new,
        fieldnames=output_fields,
        extrasaction='ignore')

    writer.writeheader()
    for row in rows:
        writer.writerow(row)

