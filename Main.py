import openpyxl
import sys
import datetime

dt_now = datetime.datetime.now()

args = sys.argv
HEAD_LINE = 1
wb = openpyxl.load_workbook(args[1], data_only=True)
ws_list = wb.worksheets
ws = wb.worksheets[0]
headerStr = ''
record = ''
tableName = ws.title
fileName=f'output/{tableName}_{dt_now.strftime('%Y%m%d%H%M%S')}_insert.sql'
f = open(fileName, 'a')
for row in ws.rows:
    recordList = []
    for cell in row:
        if(HEAD_LINE == cell.row):
            recordList.append(str(cell.value))
        else:
            recordList.append(f'\'{str(cell.value)}\'')
    if(HEAD_LINE == cell.row):
        headerStr = f'INSERT INTO {tableName} ({','.join(recordList)}) VALUES ('
    else:
        record = f'{headerStr}{','.join(recordList)});'
        f.write(record+'\n')
