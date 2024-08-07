Attribute VB_Name = "CommonAdmin"
'bu modulda butun sehifelerin acilib baglanmasi le bagli metodlar movcuddur

Public Sub OpenAllSheets()
Attribute OpenAllSheets.VB_ProcData.VB_Invoke_Func = "O\n14"
'butun sehifeleri acir
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets: ws.Visible = xlSheetVisible: Next ws
End Sub

Public Sub CloseAllSheets(shts As Variant)
'Call CloseAllSheets(Array("Home", "Main")) - bu sekilde cagirmaq olar
'gonderilen sehifelerden basqa butun sehifeleri baglayir

Dim ws As Worksheet, i As Long, flag As Boolean
flag = False

Call OpenAllSheets

'main
For Each ws In ThisWorkbook.Worksheets
  For i = LBound(shts) To UBound(shts)
    If shts(i) = ws.Name Then flag = True: Exit For
  Next i
  If flag = False Then ws.Visible = xlSheetVeryHidden
  flag = False
Next ws

End Sub

Public Sub SetPasswordAllSheets(ByVal pass As String)
'butun sehifelere password qoyur
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
  ws.Protect pass, AllowFiltering:=True
Next ws
End Sub
