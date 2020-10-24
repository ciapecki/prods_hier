# coding: utf-8
from __future__ import unicode_literals

# useful for debug
from apso_utils import msgbox

def get_color(red, green, blue):
    color = (red << 16) + (green << 8) + blue
    return color

msgbox(datetime.today().strftime("%Y-%m"))

# 1) Get values from sheet
#desktop = XSCRIPTCONTEXT.getDesktop()
#model = desktop.getCurrentComponent()
#sheet = model.CurrentController.ActiveSheet

doc = XSCRIPTCONTEXT.getDocument()
sheet = doc.getCurrentController().getActiveSheet()
# msgbox(sheet.Name)

doc.Sheets.insertNewByName('hierarchy', 1)

sheet_out = doc.Sheets['hierarchy']
sheet_out.TabColor = get_color(255, 0, 0)

#copy whole row from current to out
RngAddr = sheet.getCellRangeByName("A1:G1").getRangeAddress() 

cell = sheet_out.getCellRangeByName('A1')
cellAddr = cell.CellAddress #use this as the upper left of the paste range

sheet.copyRange(cellAddr, RngAddr)

# Get Boundaries
# columns rows

cursor = sheet.createCursor()
cursor.gotoEndOfUsedArea(False)
cursor.gotoStartOfUsedArea(True)

max_rows = cursor.Rows.Count
max_cols = cursor.Columns.Count

output_r = 1
output_latest = ''

sheet_latest = ['','','','','','','']  # holds latest written values

#for row in cursor.Rows:
#    for col in row.Columns:
#        msgbox(col.getCellByPosition(0,0).String)

#msgbox(max_cols)

for idx, value in enumerate(cursor.Rows):
    for col_idx, col_value in enumerate(value.Columns):
        if col_idx > max_cols - 1:
            break
        element = sheet.getCellByPosition(col_idx,idx+1).getString()

        #msgbox(col_idx)
        if sheet_latest[col_idx] != element:
            out_location = sheet_out.getCellByPosition(col_idx,output_r).setString(element)
            sheet_latest[col_idx] = element
            output_r += 1


# color header
#sheet_out.getCellByPosition(2,2).CellBackColor = 0
sheet_out.getCellRangeByName("A1:G1").CellBackColor = 0x038139
sheet_out.getCellRangeByName("A1:G1").CharColor = 0xFFFFFF
#sheet_out.getCellRangeByName("A1:G1").Color = 0xFFFFFF
#sheet_out.getCellByPosition(2,2).CharColor = 0x038139
#sheet_out.getCellByPosition(2,2).setString("0xFFFFFF")


# resize
sheet_out.getColumns().Width = 2000
sheet_out.getColumns().getByName("G").Width = 10000
#sheet_out.freeze_panes(1,1)

XSCRIPTCONTEXT.getDocument().getCurrentController().getActiveSheet().freeze_panes(1,1)
