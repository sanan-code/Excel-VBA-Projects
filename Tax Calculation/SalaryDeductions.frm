VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SalaryDeductions 
   Caption         =   "Calculator"
   ClientHeight    =   9492.001
   ClientLeft      =   84
   ClientTop       =   372
   ClientWidth     =   35784
   OleObjectBlob   =   "SalaryDeductions.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SalaryDeductions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'=======================================================ISCILER TEREFINDEN (Gross -> Net)=======================================================

Private Sub CommandButton1_Click()

If yoxlama = False Then Exit Sub 'yoxlama
Call mainPro 'main

End Sub

Private Sub mainPro()

Dim vcol As Double
vcol = mainPro_celbOlunanMebleg
tb_celbolunanmebleg.Value = vcol

'main
tb_gelirvergisi.Value = Round(gelirvergisi(vcol), 2)
tb_dsmf.Value = Round(dsmf(CDbl(tb_heh.Value)), 2)
tb_issizlik.Value = Round(issizlik(CDbl(tb_heh.Value)), 2)
tb_icbari.Value = Round(icbari(CDbl(tb_heh.Value)), 2)
tb_hemkar.Value = Round(hemkarlar(CDbl(tb_heh.Value)), 2)

'net
tb_net.Value = convertDotToComma(tb_heh.Value) - _
               (convertDotToComma(tb_gelirvergisi.Value) + _
               convertDotToComma(tb_dsmf.Value) + _
               convertDotToComma(tb_issizlik.Value) + _
               convertDotToComma(tb_icbari.Value) + _
               convertDotToComma(tb_hemkar.Value))

End Sub


'vergi punktlari ------------------------------------------
Private Function gelirvergisi(ByVal vcol As Double) As Double

Dim result As Double
If vcol = 0 Then GoTo cix

If opt_ds.Value Then 'dovlet
  If vcol <= 2500 Then
    If vcol <= 200 Then
      result = 0
    Else
      result = (vcol - 200) * 0.14
    End If
  Else
    result = 350 + ((vcol - 2500) * 0.25)
  End If
ElseIf opt_qds.Value Then 'qeyri dovlet
  If vcol > 8000 Then result = (vcol - 8000) * 0.14
End If

cix:
gelirvergisi = result
End Function

Private Function dsmf(ByVal vcol As Double) As Double

Dim result As Double

If opt_ds.Value Then 'dovlet
  result = vcol * 0.03
ElseIf opt_qds.Value Then 'qeyri dovlet
  If vcol <= 200 Then
    result = vcol * 0.03
  Else
    result = 6 + ((vcol - 200) * 0.1)
  End If
End If

dsmf = result
End Function

Private Function issizlik(ByVal vcol As Double) As Double
issizlik = vcol * 0.005
End Function

Private Function icbari(ByVal vcol As Double) As Double

Dim result As Double

If vcol <= 8000 Then
  result = vcol * 0.02
Else
  result = 160 + ((vcol - 8000) * 0.005)
End If

icbari = result
End Function

Private Function hemkarlar(ByVal vcol As Double) As Double
If tb_htuh.Value = "" Then Exit Function
hemkarlar = (vcol * CDbl(convertDotToComma(tb_htuh.Value))) / 100
End Function



Private Function mainPro_celbOlunanMebleg() As Double

Dim result As Double, he As Double, vm As Double

'hesablanmis emek haqqi
he = tb_heh.Value

'v.m. uzre tutulmalar
If OptionButton1.Value Then vm = 800
If OptionButton2.Value Then vm = 400
If OptionButton3.Value Then vm = 200
If OptionButton4.Value Then vm = 100
If OptionButton5.Value Then vm = 50

'final
If he <= vm Then
  result = 0
Else
  result = he - vm
End If

mainPro_celbOlunanMebleg = result
End Function




Private Function yoxlama() As Boolean

Dim result As Boolean
result = True

'hesablanan emek haqqi
If tb_heh.Value = "" Then
  MsgBox "Zehmet olmasa hesablanmis emek haqqini daxil edin"
  result = False
  GoTo cix
End If

'emek haqqi duzgun daxil edilibmi
If Not IsNumeric(tb_heh.Value) Then
  MsgBox "Zehmet olmasa hesablanmis emek haqqini duzgun daxil edilibmi"
  result = False
  GoTo cix
End If

'dovlet qeyri dovlet secimi edilibmi
If Not opt_ds.Value And Not opt_qds.Value Then
  MsgBox "Zehmet olmasa dovlet ve ya qeyri-dovlet secimi edin"
  result = False
  GoTo cix
End If

cix:
yoxlama = result
End Function










'=======================================================ISEGOTUREN TEREFINDEN (Gross -> Net)=======================================================

Private Sub CommandButton4_Click()

If yoxlama2 = False Then Exit Sub 'yoxlama
Call mainPro2 'main

End Sub

Private Sub mainPro2()

tb_ig_dsmf.Value = Round(mainPro2_dsmf(CDbl(tb_ig_heh.Value)), 2)
tb_ig_issiz.Value = Round(mainPro2_issizlik(CDbl(tb_ig_heh.Value)), 2)
tb_ig_icbari.Value = Round(mainPro2_iscbari(CDbl(tb_ig_heh.Value)), 2)

tb_ig_odenecekmebleg.Value = convertDotToComma(tb_ig_heh.Value) + convertDotToComma(tb_ig_dsmf.Value) + _
                             convertDotToComma(tb_ig_issiz.Value) + convertDotToComma(tb_ig_icbari.Value)

End Sub

'vergi punktlari -------------------------------------------------
Private Function mainPro2_dsmf(ByVal heh As Double) As Double
Dim result As Double

If opt_ig_ds.Value = True Then 'dovlet
  result = heh * 0.22
Else 'qeyri dovlet
  If heh <= 200 Then
    result = heh * 0.22
  Else
    result = 44 + ((heh - 200) * 0.15)
  End If
End If

mainPro2_dsmf = result
End Function

Private Function mainPro2_issizlik(ByVal heh As Double) As Double
Dim result As Double
result = heh * 0.005
mainPro2_issizlik = result
End Function

Private Function mainPro2_iscbari(ByVal heh As Double) As Double
Dim result As Double

If heh <= 8000 Then
  result = heh * 0.02
Else
  result = 160 + ((heh - 8000) * 0.005)
End If

mainPro2_iscbari = result
End Function


Private Function yoxlama2() As Boolean

Dim result As Boolean
result = True

'hesablanmis emek haqqi daxil edilibmi
If tb_ig_heh.Value = "" Then
  MsgBox "Zehmet olmasa meblegi duzgun daxil edin"
  result = False
  GoTo cix
End If

'daxil edilen mebleg duzgun daxil edilibmi
If Not IsNumeric(tb_ig_heh.Value) Then
  MsgBox "Zehmet olmasa meblegi duzgun daxil edin"
  result = False
  GoTo cix
End If

'dovlet qeyri dovlet secimi edilibmi
If opt_ig_ds.Value = False And opt_ig_qds.Value = False Then
  MsgBox "Zehmet olmasa dovlet ve ya qeyri dovlet secimi edin"
  result = False
  GoTo cix
End If

cix:
yoxlama2 = result
End Function



















'=======================================================ISCI TEREFINDEN (Net -> Gross)=======================================================
Private Sub CommandButton9_Click()

'yoxlama
If yoxlama3 = False Then Exit Sub
MsgBox "Hesablama bir nece saniye davam edecek. Bu muddet erzinde komputere toxnumamaginiz tovsiyye olunur", vbInformation

Dim gelirvergisi As Double, dsmf As Double, issizlik As Double, icbari As Double, hemkar As Double
Dim finalgross As Double
Dim lc As Long 'dongu sayi
Dim net As Double, gross As Double, ua As Long
Dim result As Double 'net

net = convertDotToComma(tb_isci_ng_net.Value)
gross = net * 2

'main
Do
  gelirvergisi = isci_ng_gelirVergisi(isci_ng_celbOlunanMebleg(gross))
  dsmf = isci_ng_dsmf(gross)
  issizlik = isci_ng_issizlik(gross)
  icbari = isci_ng_icbari(gross)
  hemkar = isci_ng_hemkarlar(gross)

  result = convertDotToComma(gross) - _
           convertDotToComma(gelirvergisi) - _
           convertDotToComma(dsmf) - _
           convertDotToComma(issizlik) - _
           convertDotToComma(icbari) - _
           convertDotToComma(hemkar)
  
  'yoxlama
  '1 - konkret beraber olma
  If Round(result, 3) = Round(net, 3) Then finalgross = Round(gross, 3): GoTo leave
  '2 - ?
  
  'final
  gross = gross - 0.001
  lc = lc + 1
  
  'exception
  If lc > (net / 0.001) Then Exit Do
Loop While convertDotToComma(Round(gross, 3)) <> convertDotToComma(Round(net, 3))

'final
leave:
tb_isci_ng_gelirvergisi.Value = Round(gelirvergisi, 2)
tb_isci_ng_dsmf.Value = Round(dsmf, 2)
tb_isci_ng_issizlik.Value = Round(issizlik, 2)
tb_isci_ng_icbari.Value = Round(icbari, 2)
tb_isci_ng_finalgross.Value = Round(finalgross, 2) + 0.01

End Sub



'vergi punktlari ------------------------------------------
Private Function isci_ng_gelirVergisi(ByVal vcol As Double) As Double

Dim result As Double
If vcol = 0 Then GoTo cix

If opt_isci_ng_dovlet.Value Then 'dovlet
  If vcol <= 2500 Then
    If vcol <= 200 Then
      result = 0
    Else
      result = (vcol - 200) * 0.14
    End If
  Else
    result = 350 + ((vcol - 2500) * 0.25)
  End If
ElseIf opt_isci_ng_qdovlet.Value Then 'qeyri dovlet
  If vcol > 8000 Then result = (vcol - 8000) * 0.14
End If

cix:
isci_ng_gelirVergisi = result
End Function

Private Function isci_ng_dsmf(ByVal vcol As Double) As Double

Dim result As Double

If opt_isci_ng_dovlet.Value Then 'dovlet
  result = vcol * 0.03
ElseIf opt_isci_ng_qdovlet.Value Then 'qeyri dovlet
  If vcol <= 200 Then
    result = vcol * 0.03
  Else
    result = 6 + ((vcol - 200) * 0.1)
  End If
End If

isci_ng_dsmf = result
End Function

Private Function isci_ng_issizlik(ByVal vcol As Double) As Double
isci_ng_issizlik = vcol * 0.005
End Function

Private Function isci_ng_icbari(ByVal vcol As Double) As Double

Dim result As Double

If vcol <= 8000 Then
  result = vcol * 0.02
Else
  result = 160 + ((vcol - 8000) * 0.005)
End If

isci_ng_icbari = result
End Function

Private Function isci_ng_hemkarlar(ByVal vcol As Double) As Double
If tb_isci_ng_hemkar.Value = "" Then Exit Function
isci_ng_hemkarlar = (vcol * CDbl(convertDotToComma(tb_isci_ng_hemkar.Value))) / 100
End Function


Private Function isci_ng_celbOlunanMebleg(ByVal gross As Double) As Double

Dim result As Double, he As Double, vm As Double

'hesablanmis emek haqqi
he = gross

'v.m. uzre tutulmalar
If OptionButton6.Value Then vm = 800
If OptionButton7.Value Then vm = 400
If OptionButton8.Value Then vm = 200
If OptionButton9.Value Then vm = 100
If OptionButton10.Value Then vm = 50

'final
If he <= vm Then
  result = 0
Else
  result = he - vm
End If

isci_ng_celbOlunanMebleg = result
End Function

Private Function yoxlama3() As Boolean
Dim result As Boolean
result = True

'net emek haqqi qeyd edilibmi
If tb_isci_ng_net.Value = "" Then
  MsgBox "Zehmet olmasa net emek haqqini qeyd edesiniz"
  result = False
  GoTo cix
End If

'net emek haqqi duzgun sekilde qeyd edilibmi
If Not IsNumeric(tb_isci_ng_net.Value) Then
  MsgBox "Zehmet olmasa net emek haqini duzgun qeyd edesiniz"
  result = False
  GoTo cix
End If

'dovlet ve ya qeyri dovlet secimi edilibmi
If opt_isci_ng_dovlet.Value = False And opt_isci_ng_qdovlet.Value = False Then
  MsgBox "Zehmet olmasa dovlet ve ya qeyri dovlet secimi edin"
  result = False
  GoTo cix
End If

cix:
yoxlama3 = result
End Function











'=======================================================ISEGOTUREN TEREFINDEN (Net -> Gross)=======================================================
Private Sub CommandButton11_Click()

'yoxlama
If yoxlama4 = False Then Exit Sub
MsgBox "Hesablama bir nece saniye davam edecek. Bu muddet erzinde komputere toxnumamaginiz tovsiyye olunur", vbInformation

Dim lc As Long
Dim dsmf As Double, issizlik As Double, icbari As Double
Dim heh As Double, om As Double, result As Double
om = convertDotToComma(tb_ig_ng_om.Value)
heh = om / 2

'main
Do
  dsmf = ig_ng_dsmf(heh)
  issizlik = ig_ng_issizlik(heh)
  icbari = ig_ng_iscbari(heh)

  result = convertDotToComma(dsmf) + convertDotToComma(issizlik) + convertDotToComma(icbari) + heh 'yeni odenilecek mebleg

  If Round(result, 3) = Round(om, 3) Then Exit Do
  heh = heh + 0.001
  lc = lc + 1
  
  'exception
  If lc > (om / 0.001) Then Exit Do
Loop

'final
tb_ig_ng_dsmf.Value = Round(dsmf, 2)
tb_ig_ng_issizlik.Value = Round(issizlik, 2)
tb_ig_ng_icbari.Value = Round(icbari, 2)
tb_ig_ng_heh.Value = Round(heh, 2) + 0.01

End Sub

Private Function ig_ng_dsmf(ByVal heh As Double) As Double
Dim result As Double

If opt_ig_ng_dovlet.Value = True Then 'dovlet
  result = heh * 0.22
Else 'qeyri dovlet
  If heh <= 200 Then
    result = heh * 0.22
  Else
    result = 44 + ((heh - 200) * 0.15)
  End If
End If

ig_ng_dsmf = result
End Function

Private Function ig_ng_issizlik(ByVal heh As Double) As Double
Dim result As Double
result = heh * 0.005
ig_ng_issizlik = result
End Function

Private Function ig_ng_iscbari(ByVal heh As Double) As Double
Dim result As Double

If heh <= 8000 Then
  result = heh * 0.02
Else
  result = 160 + ((heh - 8000) * 0.005)
End If

ig_ng_iscbari = result
End Function



Private Function yoxlama4() As Boolean
Dim result As Boolean
result = True

'net emek haqqi qeyd edilibmi
If tb_ig_ng_om.Value = "" Then
  MsgBox "Zehmet olmasa net emek haqqini qeyd edesiniz"
  result = False
  GoTo cix
End If

'net emek haqqi duzgun sekilde qeyd edilibmi
If Not IsNumeric(tb_ig_ng_om.Value) Then
  MsgBox "Zehmet olmasa net emek haqini duzgun qeyd edesiniz"
  result = False
  GoTo cix
End If

'dovlet ve ya qeyri dovlet secimi edilibmi
If opt_ig_ng_dovlet.Value = False And opt_ig_ng_qdovlet.Value = False Then
  MsgBox "Zehmet olmasa dovlet ve ya qeyri dovlet secimi edin"
  result = False
  GoTo cix
End If

cix:
yoxlama4 = result
End Function












'--------------------------------USERFORM--------------------------------
Private Sub CommandButton2_Click()
SpinButton1.Value = SpinButton1.Min
tb_htuh.Value = ""
End Sub

Private Sub CommandButton7_Click()
SpinButton2.Value = SpinButton2.Min
tb_isci_ng_hemkar.Value = ""
End Sub

Private Sub CommandButton3_Click()
OptionButton1.Value = False
OptionButton2.Value = False
OptionButton3.Value = False
OptionButton4.Value = False
OptionButton5.Value = False
End Sub

Private Sub CommandButton8_Click()
OptionButton6.Value = False
OptionButton7.Value = False
OptionButton8.Value = False
OptionButton9.Value = False
OptionButton10.Value = False
End Sub

Private Sub SpinButton1_Change()
tb_htuh.Value = SpinButton1.Value * 0.5
End Sub

Private Sub SpinButton2_Change()
tb_isci_ng_hemkar.Value = SpinButton2.Value * 0.5
End Sub

Private Sub ToggleButton1_Click()
If ToggleButton1.Value = True Then
  Frame1.Height = 162
  ToggleButton1.Caption = "Bagla"
Else
  Frame1.Height = 36
  ToggleButton1.Caption = "Aç"
End If
End Sub

Private Sub ToggleButton2_Click(): Call toggle_frame: End Sub
Private Sub ToggleButton3_Click(): Call toggle_frame: End Sub





Private Sub ToggleButton4_Click()
If ToggleButton4.Value = True Then
  Frame6.Height = 162
  ToggleButton4.Caption = "Bagla"
Else
  Frame6.Height = 36
  ToggleButton4.Caption = "Aç"
End If
End Sub

Private Sub UserForm_Activate()

Me.Height = 426
Me.Width = 318

Frame2.Left = 6
ToggleButton2.Caption = Label14
ToggleButton3.Caption = "Net -> Gross"

End Sub

Private Sub toggle_frame()
'toggle2 - isci, isegoturen
'toggle3 - gross->net, net->gross

'Label13 - isci
'Label14 - isegoturen

'frame2 - isciler uzre gross->net
'frame3 - isegoturen uzre gross->net
'frame4 - isciler uzre net->gross
'frame5 - isegoturen uzre net->gross
Frame2.Left = 360
Frame3.Left = 666
Frame4.Left = 990
Frame5.Left = 1314

'Default:
'isciler gross -> net

'Main

  'isciler uzre gross -> net
  If ToggleButton2.Value = False And ToggleButton3.Value = False Then
    Frame2.Left = 6
    ToggleButton2.Caption = Label14.Caption
    ToggleButton3.Caption = "Net -> Gross"
  End If
  
  'isciler uzre net -> gross
  If ToggleButton2.Value = False And ToggleButton3.Value = True Then
    Frame4.Left = 6
    ToggleButton2.Caption = Label14.Caption
    ToggleButton3.Caption = "Gross -> Net"
  End If
  
  'isegoturen uzre gross -> net
  If ToggleButton2.Value = True And ToggleButton3.Value = False Then
    Frame3.Left = 6
    ToggleButton2.Caption = Label13.Caption
    ToggleButton3.Caption = "Net -> Gross"
  End If
  
  'isegoturen uzre net -> gross
  If ToggleButton2.Value = True And ToggleButton3.Value = True Then
    Frame5.Left = 6
    ToggleButton2.Caption = Label13.Caption
    ToggleButton3.Caption = "Gross -> Net"
  End If

End Sub























'--------------------------------EXPORT--------------------------------

Private Sub CommandButton6_Click()
'isciler ucun cixaris et

Dim sr As Range, sht As String

'Yoxlama
  On Error GoTo leave

  'secim edilibmi
  If re_iscilerucun.Value = "" Then
    MsgBox "Zehmet olmasa adres secin"
    Exit Sub
  End If
  
  'setir sutun sayi bir dene olmalidir
  If Range(re_iscilerucun.Value).Rows.Count > 1 Or _
     Range(re_iscilerucun.Value).Columns.Count > 1 Then
     MsgBox "Secilen adres bir xanadan ibaret olmalidir"
     Exit Sub
  End If
    
'Main
Set sr = Range(re_iscilerucun.Value)
sht = sr.Parent.Name

With ThisWorkbook.Worksheets(sht)
  'sutun basligi
  With .Range(.Cells(sr.Row, sr.Column), .Cells(sr.Row, sr.Column + 1))
    .Merge
    .Interior.Color = vbGreen
    .HorizontalAlignment = xlCenter
    .Value = "Results"
    .Font.Bold = True
  End With
  
  .Cells(sr.Row + 1, sr.Column).Value = "Hesablanmis emek haqqi"
  .Cells(sr.Row + 2, sr.Column).Value = "V.M. tutulmalari"
  .Cells(sr.Row + 3, sr.Column).Value = "Vergiye celb olunan mebleg"
  .Cells(sr.Row + 4, sr.Column).Value = "Gelir vergisi"
  .Cells(sr.Row + 5, sr.Column).Value = "DSMF"
  .Cells(sr.Row + 6, sr.Column).Value = "Issizlik"
  .Cells(sr.Row + 7, sr.Column).Value = "Icbari"
  .Cells(sr.Row + 8, sr.Column).Value = "Hemkarlar"
  .Cells(sr.Row + 9, sr.Column).Value = "Net"
  
  .Cells(sr.Row + 1, sr.Column + 1).Value = tb_heh.Value
  .Cells(sr.Row + 2, sr.Column + 1).Value = CDbl(tb_heh.Value) - CDbl(tb_celbolunanmebleg.Value)
  .Cells(sr.Row + 3, sr.Column + 1).Value = tb_celbolunanmebleg.Value
  .Cells(sr.Row + 4, sr.Column + 1).Value = tb_gelirvergisi.Value
  .Cells(sr.Row + 5, sr.Column + 1).Value = tb_dsmf.Value
  .Cells(sr.Row + 6, sr.Column + 1).Value = tb_issizlik.Value
  .Cells(sr.Row + 7, sr.Column + 1).Value = tb_icbari.Value
  .Cells(sr.Row + 8, sr.Column + 1).Value = tb_hemkar.Value
  .Cells(sr.Row + 9, sr.Column + 1).Value = tb_net.Value
  
  .Range(.Cells(sr.Row + 9, sr.Column), .Cells(sr.Row + 9, sr.Column + 1)).Interior.Color = vbGreen
  .Cells(sr.Row + 9, sr.Column + 1).Font.Bold = True
End With

leave:
If Err.Number = 1004 Then
  MsgBox "Zehmet olmasa adresi duzgun daxil edin"
  Exit Sub
End If
End Sub


Private Sub CommandButton5_Click()
'ise goturen ucun cixaris et

Dim sr As Range, sht As String

'Yoxlama
  
  On Error GoTo leave

  'secim edilibmi
  If re_isegoturen.Value = "" Then
    MsgBox "Zehmet olmasa adres secin"
    Exit Sub
  End If
  
  'setir sutun sayi bir dene olmalidir
  If Range(re_isegoturen.Value).Rows.Count > 1 Or _
     Range(re_isegoturen.Value).Columns.Count > 1 Then
     MsgBox "Secilen adres bir xanadan ibaret olmalidir"
     Exit Sub
  End If

'Main
Set sr = Range(re_isegoturen.Value)
sht = sr.Parent.Name

With ThisWorkbook.Worksheets(sht)
  'sutun basligi
  With .Range(.Cells(sr.Row, sr.Column), .Cells(sr.Row, sr.Column + 1))
    .Merge
    .Interior.Color = vbGreen
    .HorizontalAlignment = xlCenter
    .Value = "Results"
    .Font.Bold = True
  End With
  
  .Cells(sr.Row + 1, sr.Column).Value = "Hesablanmis emek haqqi"
  .Cells(sr.Row + 2, sr.Column).Value = "DSMF"
  .Cells(sr.Row + 3, sr.Column).Value = "Issizlik"
  .Cells(sr.Row + 4, sr.Column).Value = "Icbari"
  .Cells(sr.Row + 5, sr.Column).Value = "Odenilecek mebleg"
  
  .Cells(sr.Row + 1, sr.Column + 1).Value = tb_ig_heh.Value
  .Cells(sr.Row + 2, sr.Column + 1).Value = tb_ig_dsmf.Value
  .Cells(sr.Row + 3, sr.Column + 1).Value = tb_ig_issiz.Value
  .Cells(sr.Row + 4, sr.Column + 1).Value = tb_ig_icbari.Value
  .Cells(sr.Row + 5, sr.Column + 1).Value = tb_ig_odenecekmebleg.Value
  
  .Range(.Cells(sr.Row + 5, sr.Column), .Cells(sr.Row + 5, sr.Column + 1)).Interior.Color = vbGreen
  .Cells(sr.Row + 5, sr.Column + 1).Font.Bold = True
End With

leave:
If Err.Number = 1004 Then
  MsgBox "Zehmet olmasa adresi duzgun daxil edin"
  Exit Sub
End If
End Sub

Private Sub CommandButton10_Click()
'isci ucun cixaris et (Net -> Gross)

Dim sr As Range, sht As String
'Yoxlama
  
  On Error GoTo leave

  'secim edilibmi
  If re_isci_ng.Value = "" Then
    MsgBox "Zehmet olmasa adres secin"
    Exit Sub
  End If
  
  'setir sutun sayi bir dene olmalidir
  If Range(re_isci_ng.Value).Rows.Count > 1 Or _
     Range(re_isci_ng.Value).Columns.Count > 1 Then
     MsgBox "Secilen adres bir xanadan ibaret olmalidir"
     Exit Sub
  End If

'Main
Set sr = Range(re_isci_ng.Value)
sht = sr.Parent.Name

With ThisWorkbook.Worksheets(sht)
  'sutun basligi
  With .Range(.Cells(sr.Row, sr.Column), .Cells(sr.Row, sr.Column + 1))
    .Merge
    .Interior.Color = vbGreen
    .HorizontalAlignment = xlCenter
    .Value = "Results"
    .Font.Bold = True
  End With
  
  .Cells(sr.Row + 1, sr.Column).Value = "Hesablanmis emek haqqi"
  .Cells(sr.Row + 2, sr.Column).Value = "Gelir vergisi"
  .Cells(sr.Row + 3, sr.Column).Value = "DSMF"
  .Cells(sr.Row + 4, sr.Column).Value = "Issizlik"
  .Cells(sr.Row + 5, sr.Column).Value = "Icbari"
  .Cells(sr.Row + 6, sr.Column).Value = "Odenilecek mebleg"
  
  .Cells(sr.Row + 1, sr.Column + 1).Value = tb_isci_ng_finalgross.Value
  .Cells(sr.Row + 2, sr.Column + 1).Value = tb_isci_ng_gelirvergisi.Value
  .Cells(sr.Row + 3, sr.Column + 1).Value = tb_isci_ng_dsmf.Value
  .Cells(sr.Row + 4, sr.Column + 1).Value = tb_isci_ng_issizlik.Value
  .Cells(sr.Row + 5, sr.Column + 1).Value = tb_isci_ng_icbari.Value
  .Cells(sr.Row + 6, sr.Column + 1).Value = tb_isci_ng_net.Value
  
  .Range(.Cells(sr.Row + 6, sr.Column), .Cells(sr.Row + 6, sr.Column + 1)).Interior.Color = vbGreen
  .Cells(sr.Row + 6, sr.Column + 1).Font.Bold = True
End With

leave:
If Err.Number = 1004 Then
  MsgBox "Zehmet olmasa adresi duzgun daxil edin"
  Exit Sub
End If
End Sub

Private Sub CommandButton12_Click()
'isci ucun cixaris et (Net -> Gross)

Dim sr As Range, sht As String
'Yoxlama
  
  On Error GoTo leave

  'secim edilibmi
  If re_ig_ng.Value = "" Then
    MsgBox "Zehmet olmasa adres secin"
    Exit Sub
  End If
  
  'setir sutun sayi bir dene olmalidir
  If Range(re_ig_ng.Value).Rows.Count > 1 Or _
     Range(re_ig_ng.Value).Columns.Count > 1 Then
     MsgBox "Secilen adres bir xanadan ibaret olmalidir"
     Exit Sub
  End If

'Main
Set sr = Range(re_ig_ng.Value)
sht = sr.Parent.Name

With ThisWorkbook.Worksheets(sht)
  'sutun basligi
  With .Range(.Cells(sr.Row, sr.Column), .Cells(sr.Row, sr.Column + 1))
    .Merge
    .Interior.Color = vbGreen
    .HorizontalAlignment = xlCenter
    .Value = "Results"
    .Font.Bold = True
  End With
  
  .Cells(sr.Row + 1, sr.Column).Value = "Odenilen mebleg"
  .Cells(sr.Row + 2, sr.Column).Value = "DSMF"
  .Cells(sr.Row + 3, sr.Column).Value = "Issizlik"
  .Cells(sr.Row + 4, sr.Column).Value = "Icbari"
  .Cells(sr.Row + 5, sr.Column).Value = "Hesablanan emek haqqi"
  
  .Cells(sr.Row + 1, sr.Column + 1).Value = tb_ig_ng_om.Value
  .Cells(sr.Row + 2, sr.Column + 1).Value = CDbl(tb_ig_ng_dsmf.Value)
  .Cells(sr.Row + 3, sr.Column + 1).Value = CDbl(tb_ig_ng_issizlik.Value)
  .Cells(sr.Row + 4, sr.Column + 1).Value = CDbl(tb_ig_ng_icbari.Value)
  .Cells(sr.Row + 5, sr.Column + 1).Value = CDbl(tb_ig_ng_heh.Value)
  
  .Range(.Cells(sr.Row + 5, sr.Column), .Cells(sr.Row + 5, sr.Column + 1)).Interior.Color = vbGreen
  .Cells(sr.Row + 5, sr.Column + 1).Font.Bold = True
End With

leave:
If Err.Number = 1004 Then
  MsgBox "Zehmet olmasa adresi duzgun daxil edin"
  Exit Sub
End If
End Sub


















'--------------------------------SUB PROCEDURES & FUNCTIONS--------------------------------
Private Function convertDotToComma(ByVal t As String) As Double
Dim result As Double

For i = 1 To Len(t)
  If Mid(t, i, 1) = "." Then
     t = Replace(t, ".", ",")
     Exit For
  End If
Next i

convertDotToComma = CDbl(t)
End Function

Private Sub preserveArray(ByRef a() As Double)

Dim tempArray() As Double
ReDim tempArray(LBound(a) To UBound(a) + 1, 1 To 2)

For i = LBound(a, 1) To UBound(a, 1)
  tempArray(i, 1) = a(i, 1)
  tempArray(i, 2) = a(i, 2)
Next i

a = tempArray
End Sub










