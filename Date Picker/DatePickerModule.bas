Attribute VB_Name = "DatePickerModule"
Public dpt As String

Public wrk As String
Public ws As String
Public rngR As Long
Public rngC As Long

Public uf As MSForms.UserForm
Public c As Control

'dpt date picker result target
'1 - range
'2 - on userform



Public Sub setRangeProp(ByVal wrk_ As String, ByVal ws_ As String, ByVal rngr_ As Long, ByVal rngc_ As Long)

wrk = wrk_
ws = ws_
rngR = rngr_
rngC = rngc_
dpt = 1

End Sub

Public Sub setUf(ByVal uf_ As MSForms.UserForm, ByVal c_ As Control)

Set uf = uf_
Set c = c_
dpt = 2

End Sub

