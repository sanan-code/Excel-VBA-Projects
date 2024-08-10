Attribute VB_Name = "CommonUserform"
Public Sub fillComboFromRange(ByVal wrkb As String, ByVal shtname As String, ByVal col As Integer, ByVal sr As Long, ByVal er As Long, uf As UserForm, c As Control, ByVal duplicate As Boolean)
'sehife uzerinden melumati combobox-a elave etmek
'wrkb - hansi workbooks (string)
'shtname - hansi sehife (string)
'col - sutun sayi (integer)
'sr - basladigi setir sayi
'er - bitdiyi setir sayi (-1 gonderilse bu sub icinde hesablanacaq)
'uf - userform
'c - combobox
'duplicate - duplikat deyerleri ixtisar etsin (True/False)

Dim i As Long, j As Long, flag As Boolean
flag = False

'son setiri hesablayir
If er = -1 Then er = Workbooks(wrkb).Worksheets(shtname).Cells(Rows.Count, col).End(xlUp).Row

With Workbooks(wrkb).Worksheets(shtname)
  If duplicate Then 'duplicate ile beraber
    For i = sr To er: uf.Controls(c.Name).AddItem .Cells(i, col).Value: Next i
  End If
  
  If Not duplicate Then
    For i = sr To er
      For j = i + 1 To er
        If .Cells(i, col).Value = .Cells(j, col).Value Then
          flag = True
          Exit For
        End If
      Next j
      If Not flag Then uf.Controls(c.Name).AddItem .Cells(i, col).Value
      flag = False
    Next i
  End If
End With

End Sub

Public Sub fillListFromRange(ByVal wrkb As String, _
                             ByVal sht As String, _
                             col As Variant, _
                             ByVal sr As Long, _
                             ByVal er As Long, _
                             uf As UserForm, _
                             l As Control)
'wrkb - workbookun adi
'sht - sehifenin adi
'col() - sutun nomreleri
'sr - setir nomresi baslama
'er - setir nomresi bitme (-1 olduqda sub0in ozu hesablayir)
'uf - userform
'l - list

Dim i As Long, c As Integer, r As Long

uf.Controls(l.Name).ColumnCount = UBound(col) + 1
If er = -1 Then er = Workbooks(wrkb).Sheets(sht).Cells(Rows.Count, col(UBound(col) - 1)).End(xlUp).Row
c = 0
r = 0

'main
With Workbooks(wrkb).Sheets(sht)
  For i = sr To er 'setir uzre addimlama
    uf.Controls(l.Name).AddItem
    For j = LBound(col) To UBound(col) 'sutun nomrelerinde addimlama
      uf.Controls(l.Name).List(r, c) = .Cells(i, col(j)).Value
      c = c + 1
    Next j
    r = r + 1
    c = 0
  Next i
End With

End Sub

Public Sub fillComboFromList(ByVal wrkb As String, sht As String, ByVal lo As String, col As Variant, uf As UserForm, c As Control, ByVal duplicate As Boolean)
'List Object-den deyerleri combobox-a toplamaq
'wrkb - workbook adi
'sht - sehife adi
'lo - list object adi
'col - sutun adi ve ya eded
'uf - userform
'c - combobox

Dim i As Long, j As Long, flag As Boolean, lr As Long
flag = False

With Workbooks(wrkb).Worksheets(sht).ListObjects(lo)
  lr = .ListRows.Count
  
  With .ListColumns(col)
    If duplicate Then
      For i = 1 To lr: uf.Controls(c.Name).AddItem .DataBodyRange(i).Value: Next i
    End If
    If Not duplicate Then
      For i = 1 To lr
        For j = i + 1 To lr
          If .DataBodyRange(i).Value = .DataBodyRange(j).Value Then
            flag = True
            Exit For
          End If
        Next j
        If Not flag Then uf.Controls(c.Name).AddItem .DataBodyRange(i).Value
        flag = False
      Next i
    End If
  End With
End With

End Sub

Public Sub fillListFromList(ByVal wrkb As String, _
                             ByVal sht As String, _
                             ByVal lo As String, _
                             col As Variant, _
                             ByVal uf As UserForm, _
                             ByVal l As Control)
'wrkb - workbook adi
'sht - sehife adi
'lo - listobject adi
'col - sutunlar
'uf - userform
'l - list

Dim i As Long, lr As Long, r As Long, c As Long

uf.Controls(l.Name).ColumnCount = UBound(col) + 1
r = 0
c = 0

'main
With Workbooks(wrkb).Sheets(sht).ListObjects(lo)
  lr = .ListRows.Count
  
  For i = 1 To lr 'setir nomresi
    uf.Controls(l.Name).AddItem
    For j = LBound(col) To UBound(col) 'sutunlar
      uf.Controls(l.Name).List(r, c) = .ListColumns(col(j)).DataBodyRange(i).Value
      c = c + 1
    Next j
    r = r + 1
    c = 0
  Next i
End With

End Sub
