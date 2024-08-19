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

Public Sub fillComboFromList2(ByVal wrkb As String, sht As String, ByVal lo As String, col As Variant, uf As UserForm, c As Control, ByVal duplicate As Boolean, ByVal t As String)
'List Object-den deyerleri combobox-a toplamaq (birden artiq sutunlari birlesdirir - birlesdirici vasitesi ile)
'wrkb - workbook adi
'sht - sehife adi
'lo - list object adi
'col - sutun adlarinin arrayi (Array("column1", "column2"))
'uf - userform
'c - combobox
't - birlesidrme isaresi

Dim i As Long, j As Long, k As Long, flag As Boolean, lr As Long, colName As String, colNamei As String, colNamej As String
flag = False

With Workbooks(wrkb).Worksheets(sht).ListObjects(lo)
  lr = .ListRows.Count
  
  'duplicate ile beraber
  If duplicate Then
    For i = 1 To lr 'setirler
      For j = LBound(col) To UBound(col) 'sutun arrayi
        colName = colName & .ListColumns(col(j)).DataBodyRange(i).Value & t
      Next j
      
      uf.Controls(c.Name).AddItem Trim(colName)
      colName = ""
    Next i
  End If
    
  'duplicate-siz
  If Not duplicate Then
    For i = 1 To lr 'setirler
      For j = i + 1 To lr + 1 'setirler
      
        colNamei = ""
        colNamej = ""
        
        For k = LBound(col) To UBound(col) 'sutun arrayi
          colNamei = colNamei & .ListColumns(col(k)).DataBodyRange(i).Value & t
          colNamej = colNamej & .ListColumns(col(k)).DataBodyRange(j).Value & t
        Next k
      
        If Trim(colNamei) = Trim(colNamej) Then
          flag = True 'duplicate oldugunu teyin edir
          Exit For
        End If
        
      Next j
      
      If Not flag Then uf.Controls(c.Name).AddItem colNamei
      flag = False
      colNamei = ""
      colNamej = ""
      
    Next i
  End If
  
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

Public Sub fillListFromRange_withFilter(ByVal wrkb As String, _
                                        ByVal sht As String, _
                                        ByVal sr As Long, _
                                        ByVal er As Long, _
                                        ByVal uf As UserForm, _
                                        ByVal l As Control, _
                                        col As Variant, _
                                        filt As Variant)
'wrkb - workbook adi
'sht - sehife adi
'sr - setir nomresi baslama
'er - setir nomresi bitme
'uf - userform
'l - list
'col - array(1D) - sutun nomreleri
'filt - array(2D) - filter

Dim i As Long, j As Long, k As Long, r As Long, c As Long, flag As Boolean

If er = -1 Then er = Workbooks(wrkb).Sheets(sht).Cells(Rows.Count, col(UBound(col))).End(xlUp).Row
uf.Controls(l.Name).ColumnCount = UBound(col) + 1
flag = False
c = 0
r = 0

'main
With Workbooks(wrkb).Sheets(sht)
  For i = sr To er 'setir uzre addimlama
  
    For j = LBound(filt, 1) To UBound(filt, 1) 'filter sutunlari uzre addimlama (arrayda)
      If .Cells(i, filt(j, 0)).Value = filt(j, 1) Then 'arrayda olan deyerler ile rangede olan deyerleri qarsilasdirir
        flag = True
      Else
        flag = False
        Exit For
      End If
    Next j
    
    If flag Then
      uf.Controls(l.Name).AddItem
      r = uf.Controls(l.Name).ListCount - 1
      For k = LBound(col) To UBound(col)
        uf.Controls(l.Name).List(r, c) = .Cells(i, col(k)).Value
        c = c + 1
      Next k
    End If
    
    flag = False
    c = 0
  Next i
End With

End Sub

Public Function areTextBoxesAndComboBoxesEmpty(ByVal uf As UserForm, notCheck As Variant) As Boolean
'texbox-larin ve ComboBox-larin dolu olub olmadigini kontrol eden funksiya
'uf - userform
'notCheck - yoxlanilmayan textboxlarin ve comboboxlarin adlari

'cagirilma
'1 - (me, Array("tb_iseqebul", "tb_ad")) - istisna ile
'2 - (me, Array()) - butun textboxlar ve comboboxlar

Dim t As Control, i As Integer, flag As Boolean
Dim result As Boolean
result = True
flag = False

For Each t In uf.Controls
  If TypeName(t) = "TextBox" Or TypeName(t) = "ComboBox" Then 'controls-un tipini yoxlayiriq
    Debug.Print TypeName(t)
    
    'cari olaraq controls gonderilen text box adlari ile kesisirmi
    For i = LBound(notCheck) To UBound(notCheck)
      If notCheck(i) = t.Name Then
        flag = True
        Exit For
      End If
    Next i
    
    'esas yoxlama
    If Not flag Then
      If t.Value = "" Then
        result = False
        GoTo lv
      End If
    End If
    
  End If
  
  flag = False
Next t

lv:
areTextBoxesAndComboBoxesEmpty = result
'true qayidirsa butun textboxlar doludur
'false qayidirsa textboxlar icinde box olan var (istisna olan textboxlar xaric)
End Function
