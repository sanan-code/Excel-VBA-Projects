VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SalesByNewCustomers 
   Caption         =   "Sales by New Customers"
   ClientHeight    =   6180
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5592
   OleObjectBlob   =   "SalesByNewCustomers.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SalesByNewCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rng As Range
Dim rc As Long
Private sht As Worksheet
Private dr As Integer
Private cr  As Integer
Private sr As Integer




Private Sub CommandButton3_Click()

If yoxlama = False Then Exit Sub
If MsgBox("Are you sure?", vbYesNo) = vbNo Then Exit Sub
Call mainSub
Unload Me
ThisWorkbook.Save

End Sub


Private Sub mainSub()

Dim array1() As Variant
Dim customers() As String
Dim mon As Integer, yr As Integer, i As Integer, r As Long
Dim pmon As Integer, pyr As Integer, flag As Boolean
Dim emon As Integer, eyr As Integer
Dim totalSales As Double, cust As String

flag = False
r = 0
ReDim array1(0, 2)
customers = getCustomers
mon = getStartMonandYear(True)
yr = getStartMonandYear(False)
emon = getEndMonandYear(True)
eyr = getEndMonandYear(False)

'main
Do
'evvelki ay
If mon - 1 = 0 Then
  pmon = 12
  pyr = yr - 1
Else
  pmon = mon - 1
  pyr = yr
End If
  
  'main
  For i = LBound(customers) To UBound(customers)
    cust = customers(i)
    
    'customer evvelki ayda varmi?
    For j = 2 To rc
      If Month(rng.Cells(j, dr).Value) = pmon And _
         Year(rng.Cells(j, dr).Value) = pyr Then
        If rng.Cells(j, cr).Value = cust Then 'evvelki ayda varmi?
          flag = True
          Exit For
        End If
      End If
    Next j
    
    'yoxlama
    If flag = False Then 'evvelki ayda yoxdursa?
      'cari ayda hesablama et
      For j = 2 To rc
        If Month(rng.Cells(j, dr).Value) = mon And _
           Year(rng.Cells(j, dr).Value) = yr And _
           rng.Cells(j, cr).Value = cust Then
          totalSales = totalSales + rng.Cells(j, sr).Value
        End If
      Next j
    End If
    
    'arraya toplama
    If totalSales > 0 Then
      array1(r, 0) = UCase(MonthName(mon)) & " - " & yr
      array1(r, 1) = cust
      array1(r, 2) = totalSales
      Call preserveArray(array1)
      r = r + 1
      totalSales = 0
    End If
    
    flag = False
  Next i

'novbeti ay
mon = mon + 1
If mon = 13 Then
  yr = yr + 1
  mon = 1
End If

'final
If mon > emon And yr >= eyr Then Exit Do
Loop

'report
Call rprt(array1)
Call rprt2
Call rprt3

End Sub



Private Function getStartMonandYear(ByVal t As Boolean) As Integer
'true - ay
'false - il

Dim i As Integer, j As Integer, k As Integer
Dim result As Integer, flag As Boolean

'main
For j = 2 To rc 'setir uzre
  For k = j + 1 To rc
    
    If rng.Cells(j, dr).Value < rng.Cells(k, dr).Value Then
      flag = True
    Else
      flag = False
      Exit For
    End If
    
  Next k
  
  'yoxlama
  If flag = True Then
    If t = True Then result = Month(rng.Cells(j, dr).Value)
    If t = False Then result = Year(rng.Cells(j, dr).Value)
    Exit For
  End If
Next j

'final
If result = 0 Then
  If t = True Then result = Month(rng.Cells(rc, dr).Value)
  If t = False Then result = Year(rng.Cells(rc, dr).Value)
End If
getStartMonandYear = result
End Function

Private Function getEndMonandYear(ByVal t As Boolean) As Integer
'true - ay
'false - il

Dim i As Integer, j As Integer, k As Integer
Dim result As Integer, flag As Boolean

'main
For j = 2 To rc 'setir uzre
  For k = j + 1 To rc
    
    If rng.Cells(j, dr).Value > rng.Cells(k, dr).Value Then
      flag = True
    Else
      flag = False
      Exit For
    End If
    
  Next k
  
  'yoxlama
  If flag = True Then
    If t = True Then result = Month(rng.Cells(j, dr).Value)
    If t = False Then result = Year(rng.Cells(j, dr).Value)
    Exit For
  End If
Next j

'final
If result = 0 Then
  If t = True Then result = Month(rng.Cells(rc, dr).Value)
  If t = False Then result = Year(rng.Cells(rc, dr).Value)
End If
getEndMonandYear = result
End Function

Private Function getCustomers() As String()

Dim flag As Boolean, i As Long, j As Long, r As Long
Dim result() As String
ReDim result(0)
flag = False
r = 0

For i = 2 To rc
  cust = rng.Cells(i, cr).Value
  
  For j = LBound(result) To UBound(result)
    If cust = result(j) Then
      flag = True
      Exit For
    End If
  Next j
  
  If flag = False Then
    ReDim Preserve result(r)
    result(r) = cust
    r = r + 1
  End If
  flag = False
Next i

getCustomers = result
End Function

Private Sub rprt(ByRef a1())

Dim lr As Long, i As Long

With ThisWorkbook
  .Sheets.Add
  .ActiveSheet.Name = "Sales by New Customer Report"
  lr = .ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
  
  'sutunlar
  With .ActiveSheet
    .Cells(lr, 1).Value = "Date"
    .Cells(lr, 2).Value = "Customer"
    .Cells(lr, 3).Value = "Sales"
    .Range(.Cells(lr, 1), .Cells(lr, 3)).Interior.Color = vbGreen
    .Range(.Cells(lr, 1), .Cells(lr, 3)).Font.Bold = True
    lr = .Cells(Rows.Count, 1).End(xlUp).Row
  End With
  
  'melumatlar elave edilir
  With .ActiveSheet
    For i = LBound(a1) To UBound(a1)
      .Cells(lr + 1, 1).Value = a1(i, 0)
      .Cells(lr + 1, 2).Value = a1(i, 1)
      .Cells(lr + 1, 3).Value = a1(i, 2)
      lr = lr + 1
    Next i
  End With
  
  .ActiveSheet.Range("A:C").EntireColumn.AutoFit
End With

End Sub

Private Sub rprt2()

Dim flag As Boolean
Dim rprtrng As Range, lc As Long, lr2 As Long
Set rprtrng = ThisWorkbook.ActiveSheet.Range("A1").CurrentRegion
lc = ThisWorkbook.ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
flag = False

'sutunlar
With ThisWorkbook.ActiveSheet
  .Range("A1").Copy
  .Range("E1").PasteSpecial xlPasteAll
  .Range("C1").Copy
  .Range("F1").PasteSpecial xlPasteAll
  
  lr2 = .Cells(Rows.Count, 5).End(xlUp).Row
End With

'main
With ThisWorkbook.ActiveSheet
  For i = 2 To rprtrng.Rows.Count
    
    For j = 2 To lr2 'qeyd varmi
      If .Cells(j, 5).Value = rprtrng.Cells(i, 1).Value Then
        flag = True
        
        .Cells(j, 6).Value = .Cells(j, 6).Value + rprtrng.Cells(i, 3).Value
        Exit For
      End If
    Next j
    
    If flag = False Then 'yoxdursa
      .Cells(lr2 + 1, 5).Value = rprtrng.Cells(i, 1).Value
      .Cells(lr2 + 1, 6).Value = rprtrng.Cells(i, 3).Value
    End If
    
    lr2 = .Cells(Rows.Count, 5).End(xlUp).Row
    flag = False
  Next i
  
  'final
  .Range("E:F").EntireColumn.AutoFit
End With

End Sub

Private Sub rprt3()

Dim rn As Range
Set rn = ThisWorkbook.ActiveSheet.Range("E1").CurrentRegion

ThisWorkbook.ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
ActiveChart.SetSourceData Source:=rn
ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
ThisWorkbook.ActiveSheet.Range("A1").Select
End Sub





Private Function yoxlama() As Boolean

Dim result As Boolean
result = True

If tb_customerid.Value = "" Or _
   tb_dates.Value = "" Or _
   tb_sales.Value = "" Then
  MsgBox "Please fill all sections"
  result = False
  GoTo cix
End If

cix:
yoxlama = result
End Function


' Determine columns label selections ----------------------------------
Private Sub CommandButton1_Click()
' Secilen range-den sutun adlarini liste yigir

On Error GoTo leave

Dim i As Integer
Set rng = Range(RefEdit1.Value)
Set sht = ActiveSheet 'Sheets(Left(RefEdit1.Value, InStr(1, RefEdit1.Value, "!") - 1))
rc = rng.Rows.Count
list_columns.Clear

For i = 1 To rng.Columns.Count
  list_columns.AddItem rng.Cells(1, i).Value
Next i

leave:
If Err.Number = 1004 Then
  MsgBox "Please select the range correctly"
End If
End Sub


Private Sub Label4_Click()
Call makeBoldDeterminedColumns(Label4)
End Sub
Private Sub Label5_Click()
Call makeBoldDeterminedColumns(Label5)
End Sub
Private Sub Label6_Click()
Call makeBoldDeterminedColumns(Label6)
End Sub

Private Sub makeBoldDeterminedColumns(ByVal ln As Control)
Label4.Font.Bold = False
Label5.Font.Bold = False
Label6.Font.Bold = False
ln.Font.Bold = True
End Sub

Private Sub CommandButton2_Click()
'select column type from list

For i = 0 To list_columns.ListCount - 1
  If list_columns.Selected(i) Then
    
    'label uzre yoxlama
    If Label4.Font.Bold Then cr = i + 1: tb_customerid.Value = list_columns.List(i): list_columns.Selected(i) = False: Exit For
    If Label5.Font.Bold Then dr = i + 1: tb_dates.Value = list_columns.List(i): list_columns.Selected(i) = False: Exit For
    If Label6.Font.Bold Then sr = i + 1: tb_sales.Value = list_columns.List(i): list_columns.Selected(i) = False: Exit For
    
  End If
Next i

End Sub






' Questions ----------------------------------
Private Sub Label7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Dim txt As String
txt = "First click on column name to determine. Then choose from list"
MsgBox txt
End Sub


'Other
Private Sub preserveArray(ByRef a() As Variant)

Dim temp() As Variant
temp = a
ReDim a(UBound(a, 1) + 1, UBound(a, 2))

'main
For i = LBound(temp) To UBound(temp)
  a(i, 0) = temp(i, 0)
  a(i, 1) = temp(i, 1)
  a(i, 2) = temp(i, 2)
Next i
End Sub
