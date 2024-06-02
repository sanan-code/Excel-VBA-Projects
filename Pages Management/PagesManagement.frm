VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PagesManagement 
   Caption         =   "Pages"
   ClientHeight    =   5052
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6900
   OleObjectBlob   =   "PagesManagement.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PagesManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' New page ---------------------------------------------------------------------
Private Sub CommandButton2_Click()

Dim ws As Worksheet, cp As Worksheet
Set cp = ThisWorkbook.ActiveSheet

'yoxlama
If tb_np_name.Value = "" Then MsgBox "Please write name of new page": Exit Sub
For Each ws In ThisWorkbook.Worksheets
  If LCase(ws.Name) = LCase(tb_np_name.Value) Then
    MsgBox tb_np_name.Value & " - This page name is already existed"
    Exit Sub
  End If
Next ws

'main
With ThisWorkbook
  .Sheets.Add .Sheets(1)
  .ActiveSheet.Name = Trim(tb_np_name.Value)
  .Sheets(cp.Name).Select
End With

Call fillList
End Sub



' Rename page ---------------------------------------------------------------------
Private Sub CommandButton4_Click()

Dim page As String
page = getSelectedPage

'yoxlama
If tb_r_name.Value = "" Then MsgBox "Please write name": Exit Sub
For Each ws In ThisWorkbook.Worksheets
  If LCase(ws.Name) = LCase(tb_r_name.Value) Then
    MsgBox tb_r_name.Value & " - This page name is already existed"
    Exit Sub
  End If
Next ws

'main
ThisWorkbook.Sheets(page).Name = Trim(tb_r_name.Value)

Call fillList(TextBox1.Value)
Call selectInList(tb_r_name.Value)
End Sub







' Delete ---------------------------------------------------------------------
Private Sub CommandButton1_Click()

Dim page As String
page = getSelectedPage

If page = "" Then MsgBox "Please select page from list": Exit Sub

Application.DisplayAlerts = False
ThisWorkbook.Sheets(page).Delete
Application.DisplayAlerts = True

Call fillList

End Sub



' Go to ---------------------------------------------------------------------
Private Sub CommandButton3_Click()

Dim page As String
page = getSelectedPage

If page = "" Then MsgBox "Please select page from list": Exit Sub

'eger sehife hide edilibse
ThisWorkbook.Sheets(page).Visible = xlSheetVisible
ThisWorkbook.Sheets(page).Select

End Sub


































' Userform ----------------------------------------------------
Private Sub TextBox1_Change()
Call fillList(TextBox1.Value)
End Sub

Private Sub UserForm_Activate()
Call fillList
End Sub

Private Sub fillList(Optional ByVal p As String)
list_pages.Clear

Dim ws As Worksheet, r As Range, tep As Long, flag As Boolean
flag = False

'list
If p <> "" Then
  For Each ws In ThisWorkbook.Worksheets
    If InStr(1, LCase(ws.Name), LCase(p)) > 0 Then
      list_pages.AddItem ws.Name
    End If
  Next ws
Else
  For Each ws In ThisWorkbook.Worksheets
    list_pages.AddItem ws.Name
  Next ws
End If

'indicators

'total pages
Label8.Caption = "Total page count: " & Sheets.Count

'total empty pages
For Each ws In ThisWorkbook.Worksheets
  If Excel.WorksheetFunction.CountA(ws.UsedRange) = 0 Then tep = tep + 1
Next ws
Label9.Caption = "Total empty pages: " & tep

End Sub






' Up - Down page -------------------------------------------------------
Private Sub CommandButton6_Click()
  Call movePages(getUpSelectedPage, False)
End Sub

Private Sub CommandButton5_Click()
  Call movePages(getDownSelectedPage, True)
End Sub

Private Sub movePages(ByVal i As Long, ByVal t As Boolean)
't - true - after
't - false - before
Application.DisplayAlerts = False

Dim page As String, sht As String
page = getSelectedPage
sht = ActiveSheet.Name

With ThisWorkbook
  If i <> 0 Then
    If t = True Then .Sheets(page).Move after:=.Sheets(i)
    If t = False Then .Sheets(page).Move before:=.Sheets(i)
    Call fillList(TextBox1.Value)
    Call selectInList(page)
  End If
  
  .Sheets(sht).Select
End With

Application.DisplayAlerts = True
End Sub













' Additional -------------------------------------------------------
Private Function getSelectedPage() As String

Dim result As String
result = ""

For i = 0 To list_pages.ListCount - 1
  If list_pages.Selected(i) Then
    result = list_pages.List(i)
    Exit For
  End If
Next i

getSelectedPage = result
End Function

Private Function getUpSelectedPage() As Long
On Error GoTo leave

Dim result As Long

For i = 0 To list_pages.ListCount - 1
  If list_pages.Selected(i) Then
    result = ThisWorkbook.Sheets(list_pages.List(i - 1)).Index
    Exit For
  End If
Next i

leave:
If Err.Number = 381 Then result = 0
getUpSelectedPage = result
End Function

Private Function getDownSelectedPage() As Long
On Error GoTo leave

Dim result As Long

For i = 0 To list_pages.ListCount - 1
  If list_pages.Selected(i) Then
    result = ThisWorkbook.Sheets(list_pages.List(i + 1)).Index
    Exit For
  End If
Next i

leave:
If Err.Number = 381 Then result = 0
getDownSelectedPage = result
End Function

Private Sub selectInList(ByVal p As String)
Dim k As Long
For k = 0 To list_pages.ListCount - 1
  If list_pages.List(k) = p Then
    list_pages.Selected(k) = True
    Exit For
  End If
Next k
End Sub
