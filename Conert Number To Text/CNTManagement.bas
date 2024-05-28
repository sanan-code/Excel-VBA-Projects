Attribute VB_Name = "CNTManagement"
Public Function convertNumberToText(ByVal n As String, ByVal txtP As String, ByVal txtC As String, ByVal pc As Boolean)
'n gonderilen nomre
'txtP manat
'txtC qepik
'pc sonda manat qepik yazmagi tenzimleyir

Dim finalString As String, t As Long, k As Long, c As Integer, isn As Double

'yoxlama================================================================================================================================
  'eded olaraq daxil edilibmi
  On Error GoTo leave
  isn = CDbl(n)
  '2 kesr qalana qeder yuxarlaq et
  n = Round(n, 2)

'qeyd edilen ededin kesr hissesi varmi------------------------------------------------------------------------------------------------
  c = InStr(n, ",")
  If c = 0 Then c = InStr(n, ".")

'Main================================================================================================================================
  If c = 0 Then 'tam hissesi yoxdursa------------------------------------------------------------------------------------------------
    If pc = True Then
      finalString = mainPro0(n) & txtP
    Else
      finalString = mainPro0(n)
    End If
  Else 'tam hissesi varsa------------------------------------------------------------------------------------------------
    'tam ve kesr hissesinin ayrilmasi
    t = CLng(Left(n, c - 1))
    k = CLng(Right(n, Len(n) - c))
    
    If pc = True Then
      finalString = mainPro0(t) & txtP & " "
      finalString = finalString & mainPro0(k) & txtC
    Else
      finalString = mainPro0(t) & ThisWorkbook.Sheets("CNTSource").Range("B14").Value
      finalString = finalString & mainPro0(k)
    End If
  End If

'final---------------------------
  convertNumberToText = Trim(fullFormat(finalString))

leave:
If Err.Number = 13 Then MsgBox "Zehmet olmasa ededi duzgun daxil edin"
End Function




Private Function mainPro0(ByVal n As String) As String
Dim finalString As String, temp As String, ln As Byte
ln = Len(n)

  If 1 <= ln And ln <= 3 Then 'yuz
    finalString = mainPro1(CInt(n))
  End If
  If 4 <= ln And ln <= 6 Then 'min
    finalString = mainPro1(CInt(Left(n, Len(n) - 3))) & "min "
    finalString = finalString & mainPro1(CInt(Right(n, 3)))
  End If
  If 7 <= ln And ln <= 9 Then 'milyon
    finalString = mainPro1(CInt(Left(n, Len(n) - 6))) & "milyon "
      temp = mainPro1(Mid(n, Len(n) - (Len(n) - Len(Left(n, Len(n) - 6))) + 1, 3)) 'istisna
      If temp <> "" Then finalString = finalString & temp & "min "
    finalString = finalString & mainPro1(CInt(Right(n, 3)))
  End If
  If 10 <= ln And ln <= 12 Then 'milyard
    finalString = mainPro1(CInt(Left(n, Len(n) - 9))) & "milyard "
      temp = mainPro1(Mid(n, Len(n) - (Len(n) - Len(Left(n, Len(n) - 9))) + 1, 3)) 'istisna
      If temp <> "" Then finalString = finalString & temp & "milyon "
        temp = mainPro1(Mid(n, Len(n) - (Len(n) - Len(Left(n, Len(n) - 6))) + 1, 3)) 'istisna
        If temp <> "" Then finalString = finalString & temp & "min "
    finalString = finalString & mainPro1(CInt(Right(n, 3)))
  End If
  If 13 <= ln And ln <= 15 Then 'trilyon
    finalString = mainPro1(CInt(Left(n, Len(n) - 12))) & "trilyon "
      temp = mainPro1(Mid(n, Len(n) - (Len(n) - Len(Left(n, Len(n) - 12))) + 1, 3)) 'istisna
      If temp <> "" Then finalString = finalString & temp & "milyard "
        temp = mainPro1(Mid(n, Len(n) - (Len(n) - Len(Left(n, Len(n) - 9))) + 1, 3)) 'istisna
        If temp <> "" Then finalString = finalString & temp & "milyon "
          temp = mainPro1(Mid(n, Len(n) - (Len(n) - Len(Left(n, Len(n) - 6))) + 1, 3)) 'istisna
          If temp <> "" Then finalString = finalString & temp & "min "
    finalString = finalString & mainPro1(CInt(Right(n, 3)))
  End If

mainPro0 = finalString
End Function

Private Function mainPro1(ByVal n As String) As String
Dim result As String

If n = "000" Or n = "00" Or n = "0" Then GoTo leave

If Len(n) = 1 Then
  result = get_tek(n)
End If
If Len(n) = 2 Then
  result = get_onluq(n - Right(n, 1))
  result = result & get_tek(Right(n, 1))
End If
If Len(n) = 3 Then
  result = get_tek(Left(n, 1)) & "Yüz "
  result = result & get_onluq(Right(n, 2) - Right(n, 1))
  result = Replace(result, "Bir yüz", "Yüz ")
  result = result & get_tek(Right(n, 1))
End If

leave:
mainPro1 = result
End Function















Private Function get_tek(ByVal n As Integer) As String
Dim result As String

With ThisWorkbook.Sheets("CNTSource")

  If n = 1 Then result = "Bir"
  If n = 2 Then result = "Iki"
  If n = 3 Then result = .Range("B4").Value
  If n = 4 Then result = .Range("B5").Value
  If n = 5 Then result = .Range("B6").Value
  If n = 6 Then result = .Range("B7").Value
  If n = 7 Then result = "Yeddi"
  If n = 8 Then result = .Range("B8").Value
  If n = 9 Then result = "Doqquz"

End With

get_tek = result & " "
End Function

Private Function get_onluq(ByVal n As Integer) As String
Dim result As String

With ThisWorkbook.Sheets("CNTSource")

  If n = 10 Then result = "On"
  If n = 20 Then result = "Iyirmi"
  If n = 30 Then result = "Otuz"
  If n = 40 Then result = .Range("B9").Value
  If n = 50 Then result = .Range("B10").Value
  If n = 60 Then result = .Range("B11").Value
  If n = 70 Then result = .Range("B12").Value
  If n = 80 Then result = .Range("B13").Value
  If n = 90 Then result = "Doxsan"

End With

get_onluq = result & " "
End Function

Private Function fullFormat(ByVal n As String) As String

'istisnalar
n = Replace(n, "Bir Yüz", "Yüz ")
n = Replace(n, "Bir min", "Min ")

'trim
n = Excel.WorksheetFunction.Trim(n)

'propercase
n = UCase(Left(n, 1)) & LCase(Right(n, Len(n) - 1))

fullFormat = n
End Function
