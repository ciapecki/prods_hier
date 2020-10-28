# coding: utf-8
from __future__ import unicode_literals


import uno
from com.sun.star.beans import PropertyValue

def create_instance(name, with_context=False):
    if with_context:
        instance = SM.createInstanceWithContext(name, CTX)
    else:
        instance = SM.createInstance(name)
    return instance

def call_dispatch(doc, url, args=()):
    frame = doc.getCurrentController().getFrame()
    dispatch = create_instance('com.sun.star.frame.DispatchHelper')
    dispatch.executeDispatch(frame, url, '', 0, args)
    return


def jumpFreeze(*args):

    args = ['']
    args[0] = PropertyValue()                 # Default constructor
    args[0].Name = "Nr"
    args[0].Value = 2

    call_dispatch(doc, '.uno:JumpToTable',args)
    call_dispatch(doc, '.uno:FreezePanesFirstRow')


CTX = uno.getComponentContext()
SM = CTX.getServiceManager()
doc = XSCRIPTCONTEXT.getDocument()

oDoc=XSCRIPTCONTEXT.getDocument()
oCtrl=oDoc.CurrentController
oTab=oDoc.Sheets.getByName("Products_October2020")
oCtrl.setActiveSheet(oTab)
oCtrl.freezeAtPosition(0,1)

#jumpFreeze()
