VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CNTForm 
   Caption         =   "Convert Number To Text"
   ClientHeight    =   2028
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8280.001
   OleObjectBlob   =   "CNTForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CNTForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub tb_eded_Change()
Call mainPro
End Sub

Private Sub combo_olke_Change()
Call mainPro
End Sub

Private Sub OptionButton1_Click()
Call mainPro
End Sub

Private Sub OptionButton2_Click()
Call mainPro
End Sub

Private Sub mainPro()

If tb_eded.Value = "" Then
  tb_result.Value = ""
  Exit Sub
End If

Dim n As String
Dim arg1 As String
Dim arg2 As String
Dim pc As Boolean

n = tb_eded.Value
arg1 = getCurrencyText(combo_olke.Value, 5)
arg2 = getCurrencyText(combo_olke.Value, 6)

If OptionButton1.Value = True Then pc = True
tb_result.Value = CNTManagement.convertNumberToText(n, arg1, arg2, pc)

End Sub










Private Sub UserForm_Activate()

Dim i As Integer, lr As Integer
lr = ThisWorkbook.Sheets("CNTSource").Cells(Rows.Count, 4).End(xlUp).Row

With ThisWorkbook.Sheets("CNTSource")
  For i = 3 To lr
    combo_olke.AddItem .Cells(i, 4).Value
  Next i
End With

OptionButton2.Value = True

End Sub

Private Function getCurrencyText(ByVal country As String, ByVal c As Integer) As String

Dim result As String
Dim i As Integer, lr As Integer
lr = ThisWorkbook.Sheets("CNTSource").Cells(Rows.Count, 4).End(xlUp).Row

With ThisWorkbook.Sheets("CNTSource")
  For i = 3 To lr
    If .Cells(i, 4).Value = country Then
      result = .Cells(i, c).Value
      Exit For
    End If
  Next i
End With

getCurrencyText = result
End Function
