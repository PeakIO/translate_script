from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.worksheet import Worksheet

wb: Workbook = load_workbook(filename='458041ff4b436358.xlsx')
ws: Worksheet = wb.active

res: dict = {}

for i, cells in enumerate(ws.iter_rows()):
    if i != 0:
        person_name = cells[12].value
        print(person_name)
        if person_name is not None:
            if person_name not in res:
                res[person_name] = []
            arr: list = res[person_name]
            arr.append(
                [
                    cells[4].value,
                ],
            )

for key, value in res.items():
    wbb: Workbook = Workbook()
    wss: Worksheet = wbb.active
    for v in value:
        wss.append(v)
    wbb.save(filename="./res/"+key+".xlsx")
    wbb.close()

wb.close()
