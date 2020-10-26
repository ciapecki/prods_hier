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

CTX = uno.getComponentContext()
SM = CTX.getServiceManager()
doc = XSCRIPTCONTEXT.getDocument()

args = ['']
args[0] = PropertyValue()                 # Default constructor
args[0].Name = "Nr"
args[0].Value = 2

call_dispatch(doc, '.uno:JumpToTable',args)
call_dispatch(doc, '.uno:FreezePanesFirstRow')

# FROM BASIC Macro Record
#sub freeze1row
#dim document   as object
#dim dispatcher as object
#
#rem get access to the document
#document   = ThisComponent.CurrentController.Frame
#dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
#
#dim args1(0) as new com.sun.star.beans.PropertyValue
#args1(0).Name = "Nr"
#args1(0).Value = 2
#
#dispatcher.executeDispatch(document, ".uno:JumpToTable", "", 0, args1())
#
#dispatcher.executeDispatch(document, ".uno:FreezePanesFirstRow", "", 0, Array())
#
#end sub
