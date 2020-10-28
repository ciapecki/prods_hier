from com.sun.star.sheet.ValidationType import LIST
from apso_utils import msgbox

import sys

def main():
    msgbox('ooo')
    doc = XSCRIPTCONTEXT.getDocument()
    cell = doc.Sheets[0]['B2']
    v = cell.Validation
    v.Type = LIST
    v.Formula1 = '"ONE";"TWO"'
    cell.Validation = v
    msgbox(str(v))
    return


def showSelection():
    desktop = XSCRIPTCONTEXT.getDesktop()
    selection = desktop.CurrentComponent.CurrentController.getSelection()

    msgbox(str(selection))

def showSelection2():
    doc = XSCRIPTCONTEXT.getDocument()
    selection = doc.CurrentController.Selection
    msgbox(str(selection))


def syspath():
    msgbox(str(sys.path))


g_exportedScripts = (syspath, showSelection)

