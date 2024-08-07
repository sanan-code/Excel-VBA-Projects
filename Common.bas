Attribute VB_Name = "Common"
Public Function getLastRow(ByVal shtname As String, ByVal colNum As Integer) As Long
getLastRow = ThisWorkbook.Worksheets(shtname).Cells(Rows.Count, colNum).End(xlUp).Row
End Function

Public Function getLastRow2(ByVal wrkbName As String, ByVal shtname As String, ByVal colNum As Integer) As Long
getLastRow2 = Workbooks(wrkbName).Worksheets(shtname).Cells(Rows.Count, colNum).End(xlUp).Row
End Function

Public Function getLastRow3(ByVal wrkb As String, ByVal sht As String) As Long
'gonderilen workbook ve worksheet-de mumkun en sonuncu setiri tapir

Dim final As Long
Dim result() As Variant
Dim temp As Long
Dim i As Integer
ReDim result(1 To Columns.Count)

'arraye setir nomrelerini yigit
For i = 1 To Columns.Count
  temp = Common.getLastRow2(wrkb, sht, i)
  If temp <> 1 Then result(i) = temp
Next i

'sort
Call sortArray_AtoZ(result)

'final
final = result(UBound(result))
If final = 0 Then getLastRow3 = 1 Else getLastRow3 = final
End Function

Public Function getLastColumn(ByVal shtname As String, ByVal rowNum As Integer) As Long
getLastColumn = ThisWorkbook.Worksheets(shtname).Cells(rowNum, Columns.Count).End(xlToLeft).Column
End Function

Public Function getLastColumn2(ByVal wrkbName As String, ByVal shtname As String, ByVal rowNum As Integer) As Long
getLastColumn2 = Workbooks(wrkbName).Worksheets(shtname).Cells(rowNum, Columns.Count).End(xlToLeft).Column
End Function

Public Function generateId(ByVal t As String) As String
'movuzya uygun olaraq id generate edir
generateId = t & "-" & Format(Now, "yyyymmdd-hhmmss")
End Function

Public Function areYouSure() As Boolean
Dim result As Boolean
Dim cavab As Variant
result = True
  cavab = MsgBox("Are you sure?", vbYesNo)
  If cavab = vbNo Then result = False
  areYouSure = result
End Function

Public Sub trackLog(ByVal wrkbName As String, ByVal user_ As String, ByVal oper As String, Optional ByVal note_ As String)
Dim lastrowLog As Long
lastrowLog = Common.getLastRow(wrkbName, "Logs", 1)
With ThisWorkbook.Worksheets("Logs")
  .Cells(lastrowLog + 1, 1).Value = user_ 'istifadeci
  .Cells(lastrowLog + 1, 2).Value = Format(Now, "dd.mm.yyyy - hh:mm:ss") 'date
  .Cells(lastrowLog + 1, 3).Value = oper 'oper
  If note_ <> "" Then .Cells(lastrowLog + 1, 4).Value = note_ 'note
End With
End Sub

Public Function getValueFromDB(ByVal wrkbName As String, ByVal shtname As String, ByVal lookUpValueCol As Integer, _
                              ByVal value_ As Variant, ByVal returnCol As Integer) As Variant
'wrkbName - hansi workbook-da axtarir
'shtName - hansi sehifede axtarir
'lookUpValueCol - hansi sutunda axtaris edecek
'value_ - neye gore axtarir
'returnCol - hansi sutundan return edecek
Dim result As Variant
Dim lastrow As Long
lastrow = Common.getLastRow2(wrkbName, shtname, lookUpValueCol)
With Workbooks(wrkbName).Worksheets(shtname)
  For i = 2 To lastrow
    If CVar(.Cells(i, lookUpValueCol).Value) = value_ Then
      result = .Cells(i, returnCol).Value
      Exit For
    End If
  Next i
End With
getValueFromDB = result
End Function

Public Function getValueFromDB2(ByVal wrkbName As String, ByVal shtname As String, ByVal lookUpValueCol As Integer, _
                              ByVal value_ As Variant, ByVal returnCol As Integer) As Variant()
'wrkbName - hansi workbook-da axtarir
'shtName - hansi sehifede axtarir
'lookUpValueCol - hansi sutunda axtaris edecek
'value_ - neye gore axtarir
'returnCol - hansi sutundan return edecek
Dim result() As Variant
Dim lastrow, arrayLen, r As Long
lastrow = Common.getLastRow(shtname, lookUpValueCol)
r = 0

'nece dene deyer getirilecek onu teyin edir (Array ucun) (setir uzre)
With Workbooks(wrkbName).Worksheets(shtname)
  For i = 2 To lastrow
    If CVar(.Cells(i, lookUpValueCol).Value) = value_ Then arrayLen = arrayLen + 1
  Next i
End With

'redim array
ReDim result(arrayLen - 1)

'melumatlar arraya daxil edilir
With Workbooks(wrkbName).Worksheets(shtname)
  For i = 2 To lastrow
    If CVar(.Cells(i, lookUpValueCol).Value) = value_ Then result(r) = CVar(.Cells(i, returnCol).Value): r = r + 1
  Next i
End With

getValueFromDB2 = result
End Function

Public Function getValueFromDB3(ByVal wrkbName As String, ByVal shtname As String, ByVal lookUpValueCol As Integer, _
                              ByVal value_ As Variant, returnCol() As Integer) As Variant()
'wrkbName - hansi workbook-da axtarir
'shtName - hansi sehifede axtarir
'lookUpValueCol - hansi sutunda axtaris edecek
'value_ - neye gore axtarir
'returnCol - hansi sutunlardan return edecek
Dim result() As Variant
Dim lastrow, r As Long
lastrow = Common.getLastRow(shtname, lookUpValueCol)
r = 0

'nece dene deyer getirilecek onu teyin edir (Array ucun) (setir uzre)
With Workbooks(wrkbName).Worksheets(shtname)
  For i = 2 To lastrow
    If CVar(.Cells(i, lookUpValueCol).Value) = value_ Then arrayLen = arrayLen + 1
  Next i
End With

'array nece mertebeli olacaq
arrayCol = UBound(returnCol)

'redim array
ReDim result(arrayLen - 1, arrayCol)

'melumatlar arraya daxil edilir
With Workbooks(wrkbName).Worksheets(shtname)
  For i = 2 To lastrow
    If CVar(.Cells(i, lookUpValueCol).Value) = value_ Then
    
      For j = LBound(returnCol) To UBound(returnCol)
        result(r, j) = CVar(.Cells(i, returnCol(j)).Value)
      Next j
      
      r = r + 1
    End If
  Next i
End With

getValueFromDB3 = result
End Function

Public Sub arrayToRange(ByVal wrkbName As String, ByVal shtname As String, ByVal startRange As String, _
                              ByRef mainArray() As Variant, Optional ByVal colName As String)

Dim i, startRow, startCol As Long
startRow = Range(startRange).Row
startCol = Range(startRange).Column

With Workbooks(wrkbName).Worksheets(shtname)
  
  If colName <> "" Then 'sutun basliginin baslamasi
    .Range(startRange).Value = colName
    startRow = startRow + 1
  End If
  
  'main
  For i = LBound(mainArray) To UBound(mainArray)
    .Cells(startRow, startCol).Value = mainArray(i)
    startRow = startRow + 1
  Next i
  
End With

End Sub

Public Sub multiArrayToRange(ByVal wrkbName As String, ByVal shtname As String, ByVal startRange As String, _
                              ByRef mainArray() As Variant, ByRef colNames() As String)

Dim i, startRow, startCol As Long
startRow = Range(startRange).Row
startCol = Range(startRange).Column

With Workbooks(wrkbName).Worksheets(shtname)
  
  'sutun basligi
  If UBound(colNames) > 0 Then
    For i = LBound(colNames) To UBound(colNames)
      .Cells(startRow, startCol).Value = colNames(i)
      startCol = startCol + 1
    Next i
    startRow = Range(startRange).Row + 1
    startCol = Range(startRange).Column
  End If
  
  'melumatlar elave edilir
  For i = LBound(mainArray, 1) To UBound(mainArray, 1)
    For j = LBound(mainArray, 2) To UBound(mainArray, 2)
      .Cells(startRow, startCol).Value = mainArray(i, j)
      startCol = startCol + 1
    Next j
    startRow = startRow + 1
    startCol = Range(startRange).Column
  Next i
  
End With

End Sub

Public Function getValueFromCell(ByVal wrkbName As String, ByVal shtname As String, ByVal r As Long, ByVal c As Long) As String
  getValueFromCell = Workbooks(wrkbName).Worksheets(shtname).Cells(r, c).Value
End Function

Public Sub fillEmptyCells(ByVal wrkbName As String, ByVal sheetName As String, ByVal rowNum As Integer, ByVal lastCol As Integer)
'qeyd edilen sehifede bosluqlari doldurur
'sheetName - hansi sehifede is gorecek
'hansi setirde is gorecek
'lastCol - hansi sutuna qeder davam edecek yoxlamaga
Dim rng, controlRange As Range
With Workbooks(wrkbName).Worksheets(sheetName)
  Set controlRange = .Range(.Cells(rowNum, 4), .Cells(rowNum, lastCol))
End With
  For Each rng In controlRange
    If rng.Value = "" Then rng.Value = "-"
  Next
End Sub

Public Function getColNumberOfLetter(ByVal letter As String) As Integer

Dim letterArray(1 To 16384) As String
Dim result, tempResult As String
Dim addresses_, colLetter As String

Set addresses_ = Range("A1:XFD1")
rw = 1

'array-a herfleri elave edir
For Each r In addresses_
colLetter = r.Address

  For i = 1 To Len(colLetter)
    If Not IsNumeric(Mid(colLetter, i, 1)) And _
        Mid(colLetter, i, 1) <> "$" And _
        Mid(colLetter, i, 1) <> ":" Then
        
      tempResult = tempResult & Mid(colLetter, i, 1)
    End If
  Next i
  
  letterArray(rw) = tempResult
  rw = rw + 1
  tempResult = ""
Next r

'yoxlama
For i = LBound(letterArray) To UBound(letterArray)
  If letterArray(i) = letter Then
    result = i
    Exit For
  End If
Next i

getColNumberOfLetter = result
End Function

Public Sub autoFitColumns(ByVal wrkbName As String, ByVal ws As String, ByVal startCol As String, ByVal endCol As String)
  Workbooks(wrkbName).Worksheets(ws).Range(startCol & ":" & endCol).EntireColumn.AutoFit
End Sub

Public Sub addNewRecordIfNeeded(ByVal shtname As String, ByVal newRecord As String, ByVal col As Integer)
'sutun 2 sci setirden baslayirsa bu istifade oluna biler

Dim flag As Boolean
Dim i, lastrow As Integer
lastrow = Common.getLastRow(shtname, col)
flag = False

'main
With ThisWorkbook.Worksheets(shtname)
  'yoxlama
  For i = 2 To lastrow
    If .Cells(i, col).Value = newRecord Then
      flag = True
      Exit For
    End If
  Next i

  'yoxdursa elave et
  If flag = False Then .Cells(lastrow + 1, col).Value = newRecord
End With

End Sub

Public Function sortArray_AtoZ(ByRef a() As Variant)

Dim i As Integer
Dim j As Integer
Dim temp As Integer

'sorting array from A to Z
For i = LBound(a) To UBound(a)
  For j = i + 1 To UBound(a)
    If a(i) > a(j) Then
      temp = a(j)
      a(j) = a(i)
      a(i) = temp
    End If
  Next j
Next i

sortArray_AtoZ = a
End Function

Public Sub makeProperCase(t As String)
'hem trim edir ve hemcinin proper case edir

Dim i As Integer, temp As String

'bosluqlari temizle
For i = 1 To Len(t)
  If Not (Mid(t, i, 1) = " " And Mid(t, i + 1, 1) = " ") Then temp = temp & Mid(t, i, 1)
Next i
t = Trim(temp)
temp = ""
temp = UCase(Left(t, 1))

'2 den baslayaraq proper case et
For i = 2 To Len(t)
  If Mid(t, i, 1) = " " And Mid(t, i + 1, 1) <> "" Then
    temp = temp & " "
    temp = temp & UCase(Mid(t, i + 1, 1))
    i = i + 1
  Else
    temp = temp & Mid(t, i, 1)
  End If
Next i

t = Trim(temp)
End Sub

Private Sub addValidationList(ByVal rng As Range, _
                             ByVal l As String, _
                             Optional ByVal it As String, _
                             Optional ByVal im As String, _
                             Optional ByVal et As String, _
                             Optional ByVal em As String)
'rng - hansi rangde validation olacaq
'l - data source
'it - input title
'im - input message
'et - error title
'em - error message

rng.Select
With Selection.Validation
  .Delete
  .add Type:=xlValidateList, _
       AlertStyle:=xlValidAlertStop, _
       Operator:=xlBetween, _
       Formula1:=l
  
  .IgnoreBlank = True
  .InCellDropdown = True
  .InputTitle = it
  .InputMessage = im
  .ErrorTitle = et
  .ErrorMessage = em
  .ShowInput = True
  .ShowError = True
End With
    
End Sub

Public Function getExperienceByYMD(eArray() As Date, ByVal t As Byte) As Variant
'eArray - iki olculu massiv olmalidir
't - 1 olduqda string qaytarir (1 il 1 ay 1 gun)
't - 2 olduqda integer tipli array qaytarir

Dim result As Variant
Dim temp As Long
Dim i As Byte

'main
For i = 1 To UBound(eArray, 1)
  temp = temp + (eArray(i, 2) - eArray(i, 1) + 1)
Next i

'final
  il = Int(temp / 365.2)
  ay = Int((temp - il * 365.2) / 30.4)
  gun = Int((temp - il * 365.2) - (ay * 30.4))

If t = 1 Then
  result = il & " il " & ay & " ay " & gun & " gün"
ElseIf t = 2 Then
  ReDim result(1 To 3)
  result(1) = il
  result(2) = ay
  result(3) = gun
End If

getExperienceByYMD = result
End Function

Public Function decimalToBinary(ByVal n As Long) As Long
Dim result As String, qaliq As Long
Do
  qaliq = n Mod 2
  n = Excel.WorksheetFunction.RoundDown(n / 2, 0)
  result = qaliq & result
Loop Until n = 0
decimalToBinary = CLng(result)
End Function

Public Function binaryToDecimal(ByVal n As Long) As Long
Dim result As Long, i As Long
For i = 1 To Len(CStr(n))
  result = ((2 ^ (Len(CStr(n)) - i)) * Mid(CStr(n), i, 1)) + result
Next i
binaryToDecimal = CLng(result)
End Function

Public Function areDatesIntersect(ByVal sd As Date, ByVal ed As Date, ByVal scd As Date, ByVal ecd As Date)
'sd - start date
'ed - end date
'scd - checking start date
'ecd - checking end date

Dim result As Boolean
result = False

'yoxlama
If (sd - ed > 0) Or (scd - ecd > 0) Then Exit Function

'main
If (sd <= scd And scd <= ed) Or (sd <= ecd And ecd <= ed) Then result = True

areDatesIntersect = result
End Function

Public Function isDateIntersect(ByVal sd As Date, ByVal ed As Date, ByVal d As Date) As Boolean
'sd - start date
'ed - end date
'd - checking date

Dim result As Boolean
result = False

'yoxlama
If sd - ed > 0 Then Exit Function

'main
If sd <= d And d <= ed Then isDateIntersect = True

isDateIntersect = result
End Function

Public Function areDatesIntersectDays(ByVal sd As Date, ByVal ed As Date, ByVal scd As Date, ByVal ecd As Date) As Integer
'nece gun ustu uste dusur

Dim result As Integer
result = 0

'yoxlama
If (sd - ed > 0) Or (scd - ecd > 0) Then Exit Function

'main
If (sd <= scd And ecd <= ed) Then
  result = ecd - scd + 1
  GoTo lv
End If
If (sd <= scd And scd <= ed) Then
  result = ed - scd + 1
  GoTo lv
End If
If (sd <= ecd And ecd <= ed) Then
  result = ecd - sd + 1
  GoTo lv
End If

lv:
areDatesIntersectDays = result
End Function
