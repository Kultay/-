# import xlrd
# import csv
#
# wb = xlrd.open_workbook(r'C:\Users\Danik\Downloads\КАТО_17.11.2023.xls')
# ws = wb.sheet_by_index(0)
#
# col = csv.writer(open("КАТО_old.csv", 'w', newline="", encoding="utf"))
#
# for row in range(ws.nrows):
#     col.writerow(ws.row_values(row))
#
# with open("КАТО_old.csv", 'r', newline="", encoding="utf") as csv_file:
#     reader = csv.reader(csv_file)
#     # rows = list(reader)
#     #
#     # output_fields = reader.fieldnames
#     # output_fields.remove("nn")
#     # output_fields.remove("ab")
#     # output_fields.remove("cd")
#     # output_fields.remove("ef")
#     # output_fields.remove("hij")
#     # output_fields.remove("k")
#
#     # output_fields.append("id")
#     # output_fields.append("parent_id")
#
#     # for i, row in enumerate(rows, start=1):
#     #     row['id'] = i
#
#         # if float(row["ab"]) > 0 and float(row["cd"]) > 0 and float(row["ef"]) == 0 and float(row["hij"]) == 0:
#         #     previous_i1 = i - 1
#         #     row["parent_id"] = previous_i1
#         # elif float(row["ab"]) > 0 and float(row["cd"]) > 0 and float(row["ef"]) > 0 and float(row["hij"]) == 0:
#         #     previous_i2 = i
#         #     row["parent_id"] = previous_i1 + 1
#         # elif float(row["ab"]) > 0 and float(row["cd"]) > 0 and float(row["ef"]) > 0 and float(row["hij"]) > 0:
#         #     row["parent_id"] = previous_i2
#         # else:
#         #     row["parent_id"] = None
#         #
#         # if row['te'].endswith('.0'):
#         #     row['te'] = str(int(float(row['te'])))
#
# with open('КАТО_new.csv', 'w', newline="", encoding="utf") as csv_new:
#     writer = csv.writer(
#         csv_new)
#
#
#     writer.writeheader()
#     for row in rows:
#         writer.writerow(row)
#



import xlrd
import csv

wb = xlrd.open_workbook(r'C:\Users\Danik\Downloads\КАТО_17.11.2023.xls')
ws = wb.sheet_by_index(0)

rows = []

for row_index in range(ws.nrows):
    rows.append(ws.row_values(row_index))


header = rows[0]
output_fields = [field for field in header if field not in ["ab", "cd", "ef", "hij", "k", "nn"]] + ["id", "parent_id"]


previous_i1 = None
previous_i2 = None


for i, row in enumerate(rows[1:], start=1):
    row.append(i)

    if float(row[header.index("ab")]) > 0 and float(row[header.index("cd")]) > 0 \
            and float(row[header.index("ef")]) == 0 and float(row[header.index("hij")]) == 0:
        row.append(previous_i1 )
        previous_i1 = i-1
    elif float(row[header.index("ab")]) > 0 and float(row[header.index("cd")]) > 0 \
            and float(row[header.index("ef")]) > 0 and float(row[header.index("hij")]) == 0:
        row.append(previous_i1 + 1)
        previous_i2 = i
    elif float(row[header.index("ab")]) > 0 and float(row[header.index("cd")]) > 0 \
            and float(row[header.index("ef")]) > 0 and float(row[header.index("hij")]) > 0:
        row.append(previous_i2)
    else:
        row.append(None)


    if str(row[header.index('te')]).endswith('.0'):
        row[header.index('te')] = str(int(float(row[header.index('te')])))


columns_to_remove = [header.index("ab"), header.index("cd"), header.index("ef"), header.index("hij"), header.index("k"), header.index("nn")]
rows = [[row[i] for i in range(len(row)) if i not in columns_to_remove] for row in rows]


with open('КАТО_new.csv', 'w', newline="", encoding="utf") as csv_new:
    writer = csv.writer(csv_new)
    writer.writerow(output_fields)
    writer.writerows(rows[1:])

