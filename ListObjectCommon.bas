Attribute VB_Name = "ListObjectCommon"
Public Function getListObjectColumnsName(ByVal wrkb As String, ByVal sht As String, ByVal lo As String) As String()

Dim result() As String
Dim i As Integer

With Workbooks(wrkb).Worksheets(sht).ListObjects(lo)
  For i = 1 To .HeaderRowRange.Count
    ReDim Preserve result(i - 1)
    result(i - 1) = .HeaderRowRange(i).Value
  Next i
End With

getListObjectColumnsName = result
End Function

Public Function getListObjectAValue(ByVal wrkb As String, ByVal sht As String, ByVal lo As String, ByVal lookUpCol As Variant, ByVal lookUpVal As Variant, ByVal resultCol As Variant) As Variant
'bir eded deyer qaytarir
'lookUpCol (variant) - hansi sutunda axtaracaq (hem adini gondere bilerik hem nomresini)
'lookUpVal (variant) - axtarilan ortaq deyer
'resultCol (variant) - hansi sutundan melumat gonderecek (hem adini gondere bilerik hem nomresini)

Dim result As Variant

With Workbooks(wrkb).Worksheets(sht).ListObjects(lo)
  For i = 1 To .ListRows.Count
    If .ListColumns(lookUpCol).DataBodyRange(i).Value = lookUpVal Then
      result = .ListColumns(resultCol).DataBodyRange(i).Value
      Exit For
    End If
  Next i
End With

getListObjectAValue = result
End Function

Public Function getListObjectValueList(ByVal wrkb As String, ByVal sht As String, ByVal lo As String, ByVal lookUpCol As Variant, ByVal lookUpVal As Variant, ByVal resultCol As Variant) As Variant()
'birden cox deyeri list kimi return edir
'lookUpCol (variant) - hansi sutunda axtaracaq (hem adini gondere bilerik hem nomresini)
'lookUpVal (variant) - axtarilan ortaq deyer
'resultCol (variant) - hansi sutundan melumat gonderecek (hem adini gondere bilerik hem nomresini)

Dim result() As Variant
Dim r As Long
r = 0

With Workbooks(wrkb).Worksheets(sht).ListObjects(lo)
  For i = 1 To .ListRows.Count
    If .ListColumns(lookUpCol).DataBodyRange(i).Value = lookUpVal Then
      ReDim Preserve result(r)
      result(r) = .ListColumns(resultCol).DataBodyRange(i).Value
      r = r + 1
    End If
  Next i
End With

getListObjectValueList = result
End Function

Public Function deleteRowsListObject(ByVal wrkb As String, ByVal sht As String, ByVal lo As String, ByVal lookUpCol As Variant, ByVal lookUpVal As Variant) As Long
'lookUpVal (variant) - axtarilan ortaq deyer
'resultCol (variant) - hansi sutundan melumat gonderecek (hem adini gondere bilerik hem nomresini)

Dim deleteRowsCount As Long
Dim i As Long

With Workbooks(wrkb).Worksheets(sht).ListObjects(lo)
  For i = .ListRows.Count To 1 Step -1
    If .ListColumns(lookUpCol).DataBodyRange(i).Value = lookUpVal Then
      .DataBodyRange.Rows(i).Delete
      deleteRowsCount = deleteRowsCount + 1
    End If
  Next i
End With

deleteRowsListObject = deleteRowsCount
End Function

Public Function addRowsAfterSpecifiedValue(ByVal wrkb As String, ByVal sht As String, ByVal lo As String, ByVal lookUpCol As Variant, ByVal lookUpVal As Variant, rc As Long, ba As Boolean) As Long
'lookUpVal (variant) - axtarilan ortaq deyer
'resultCol (variant) - hansi sutundan melumat gonderecek (hem adini gondere bilerik hem nomresini)
'ba - true:after, false:before

Dim addedRowsCount As Long, i As Long, j As Long, r As Long
If ba Then r = 1 Else r = 0

With Workbooks(wrkb).Worksheets(sht).ListObjects(lo)
  For i = .ListRows.Count To i Step -1
    If .ListColumns(lookUpCol).DataBodyRange(i).Value = lookUpVal Then
      For j = 1 To rc: .ListRows.Add i + r: Next j
      addedRowsCount = addedRowsCount + rc
    End If
  Next i
End With

addRowsAfterSpecifiedValue = addedRowsCount
End Function

Public Function deleteEmptyRows1(ByVal wrkb As String, ByVal sht As String, ByVal lo As String) As Long
'eger setir tam bosdursa

Dim i As Long, j As Long, deletedRowsCount As Long
Dim rc As Long, cc As Long, flag As Boolean
flag = True

With Workbooks(wrkb).Worksheets(sht).ListObjects(lo)
rc = .ListRows.Count
cc = .ListColumns.Count

  For i = rc To 1 Step -1
    For j = 1 To cc
      If Not .ListColumns(j).DataBodyRange(i).Value = "" Then
        flag = False
        Exit For
      End If
    Next j
    
    If flag Then
      .DataBodyRange.Rows(i).Delete
      deletedRowsCount = deletedRowsCount + 1
    End If
    flag = True
  Next i
End With

deleteEmptyRows1 = deletedRowsCount
End Function

Public Function deleteEmptyRows2(ByVal wrkb As String, ByVal sht As String, ByVal lo As String) As Long
'eger setirde en az bir xana bosdursa

Dim i As Long, j As Long, deletedRowsCount As Long
Dim rc As Long, cc As Long, flag As Boolean
flag = False

With Workbooks(wrkb).Worksheets(sht).ListObjects(lo)
rc = .ListRows.Count
cc = .ListColumns.Count

  For i = rc To 1 Step -1
    For j = 1 To cc
      If .ListColumns(j).DataBodyRange(i).Value = "" Then
        flag = True
        Exit For
      End If
    Next j
    
    If flag Then
      .DataBodyRange.Rows(i).Delete
      deletedRowsCount = deletedRowsCount + 1
    End If
    flag = False
  Next i
End With

deleteEmptyRows2 = deletedRowsCount
End Function

Public Function deleteEmptyRows3(ByVal wrkb As String, ByVal sht As String, ByVal lo As String, ByVal col As Variant) As Long
'eger teyin edilmis sutunda bosluq varsa

Dim i As Long, deletedRowsCount As Long

With Workbooks(wrkb).Worksheets(sht).ListObjects(lo)
  For i = .ListRows.Count To 1 Step -1
    If .ListColumns(col).DataBodyRange(i).Value = "" Then
      .DataBodyRange.Rows(i).Delete
      deletedRowsCount = deletedRowsCount + 1
    End If
  Next i
End With

deleteEmptyRows3 = deletedRowsCount
End Function















