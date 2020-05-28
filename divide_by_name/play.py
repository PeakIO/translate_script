import os
from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.worksheet import Worksheet

filename = "13"

wb: Workbook = load_workbook(filename=filename + '.xlsx')
ws: Worksheet = wb.active

res: dict = {}

for i, cells in enumerate(ws.iter_rows()):
    if i != 0:
        person_name = cells[12].value
        if person_name is not None:
            if person_name not in res:
                res[person_name] = []
            arr: list = res[person_name]
            arr.append(
                [
                    cells[4].value,
                ],
            )

if not os.path.exists(filename):
    os.makedirs(filename)

for name, contents in res.items():
    wbb: Workbook = Workbook()
    wss: Worksheet = wbb.active
    for rows in contents:
        print(rows)
        wss.append(rows)
    wbb.save(filename="./{path}/{key}.xlsx".format(path=filename, key=name))
    wbb.close()

wb.close()
