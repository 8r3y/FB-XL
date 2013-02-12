import fdb
import xlwt
import xlrd
from datetime import datetime

con = fdb.connect(
    dsn='127.0.0.1:C:/Python27/fdb/test/fbtest.fdb',
    user='sysdba', password='masterkey',
    )

cur = con.cursor()

cur.execute("select * from testspec")

wb_report = xlwt.Workbook()
wb = xlwt.Workbook()

ws_report = wb_report.add_sheet('ReportSheet')
ws = wb.add_sheet('WorkSheet')

style1 = xlwt.easyxf(num_format_str='D-MMM-YY')

ws_report.write(0, 0, datetime.now(), style1)

j=0
fieldIndices = range(len(cur.description))
print fieldIndices
for row in cur:
    for fieldIndex in fieldIndices:
        fieldValue = str(row[fieldIndex])
        ws.write(j, fieldIndex, fieldValue)
    j=j+1
wb.save('export.xls')

book = xlrd.open_workbook("export.xls")
print "The number of worksheets is", book.nsheets
print "Worksheet name(s):", book.sheet_names()
sh = book.sheet_by_index(0)
print sh.name, sh.nrows, sh.ncols

insertStatement = cur.prep("insert into IMPORTSPEC (BAR, CL_ART, CL_NAME) values (?,?,?)")

for row in range(sh.nrows):
    inputRows = []
    for col in range(sh.ncols):
        val = sh.cell_value(row, col)
        if isinstance(val, unicode):
            val = val.encode('utf8');
        inputRows.append(val)
    cur.execute(insertStatement, inputRows)

con.commit()

ws_report.write(1, 0, datetime.now(), style1)

wb_report.save('report.xls')
