
On Error Resume Next
Dim MPFController : Set MPFController = CreateObject("MPF.Controller")
MPFController.Run
MPFController.Switch("0-0-3")=1
MPFController.Switch("0-0-8")=1
MPFController.Switch("0-0-9")=1
MPFController.Switch("0-0-10")=1
MPFController.Switch("0-0-11")=1
If Err Then MsgBox "MPF Not Setup"
On Error GoTo 0
