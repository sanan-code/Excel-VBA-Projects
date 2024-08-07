Attribute VB_Name = "LogPassManagement"
Public Sub LogAndPassControl(ByVal log_ As String, ByVal pass_ As String)

Dim i As Long, j As Integer, lrLogSource As Long, lcLogSource As Integer, shtname As String
With ThisWorkbook.Sheets("LoginSource").ListObjects("LogPassTable")
  lrLogSource = .ListRows.Count
  lcLogSource = .ListColumns.Count
End With

'yoxlama
If log_ = "" Or pass_ = "" Then Exit Sub

On Error Resume Next
'main
With ThisWorkbook.Sheets("LoginSource").ListObjects("LogPassTable")
  For i = 1 To lrLogSource
    If .ListColumns("login").DataBodyRange(i).Value = log_ And .ListColumns("password").DataBodyRange(i).Value = pass_ Then
      For j = 4 To lcLogSource
        If .ListColumns(j).DataBodyRange(i).Value = "OK" Then
          shtname = .ListColumns(j).Range(1).Value
          ThisWorkbook.Sheets(shtname).Visible = xlSheetVisible
        End If
      Next j
    End If
  Next i
End With

End Sub

Public Sub DeleteUser(ByVal username As String)

Dim i As Long, lrLogSource As Long
lrLogSource = ThisWorkbook.Sheets("LoginSource").ListObjects("LogPassTable").ListRows.Count

'main
With ThisWorkbook.Sheets("LoginSource").ListObjects("LogPassTable")
  lrLogSource = .ListRows.Count
  
  For i = 1 To lrLogSource
    If .ListColumns("user").DataBodyRange(i).Value = username Then
      .ListRows(i).Delete
      Exit For
    End If
  Next i
End With

End Sub

Public Sub AddNewUser(ByVal username As String, ByVal log_ As String, ByVal pass_ As String, shts As Variant)

Dim lr As Long, i As Integer, j As Integer, lc As Integer

With ThisWorkbook.Sheets("LoginSource").ListObjects("LogPassTable")
  .ListRows.add
  lr = .ListRows.Count
  
  .ListColumns("user").DataBodyRange(lr).Value = username
  .ListColumns("login").DataBodyRange(lr).Value = log_
  .ListColumns("password").DataBodyRange(lr).Value = pass_
  
  lc = .ListColumns.Count
  For i = 4 To lc
    sht = .ListColumns(i).Range(1).Value
    For j = LBound(shts) To UBound(shts)
      If shts(j) = sht Then
        .ListColumns(sht).DataBodyRange(lr).Value = "OK"
        Exit For
      End If
    Next j
  Next i
End With

End Sub

