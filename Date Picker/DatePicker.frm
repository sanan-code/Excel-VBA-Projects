VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DatePicker 
   Caption         =   "Date Picker"
   ClientHeight    =   4140
   ClientLeft      =   84
   ClientTop       =   360
   ClientWidth     =   3588
   OleObjectBlob   =   "DatePicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private ufb As Boolean

Private Sub combo_month_Change()

If ufb = False Then Exit Sub
Call fillXdates: Call fillCurrentDates: Call formatXs: Call formatAllLabels

'lb_today
Call lb_today_

End Sub

Private Sub combo_year_Change()

If ufb = False Then Exit Sub
Call fillXdates: Call fillCurrentDates: Call formatXs: Call formatAllLabels

'lb_today
Call lb_today_

End Sub

Private Sub CommandButton1_Click()

ufb = False
combo_month.Value = getMonth(Month(Date))
combo_year.Value = Year(Date)
ufb = True

Call fillXdates: Call fillCurrentDates: Call formatXs: Call formatAllLabels: Call formatDayinActivate

lb_today.Caption = ""

End Sub










'Esas prosesler---------------------------------------------------------------------------------------------------

Private Sub UserForm_Activate()
Application.ScreenUpdating = False

Dim i As Long

'month
For i = 1 To 12: combo_month.AddItem getMonth(i): Next i
'year
For i = 2050 To 1901 Step -1: combo_year.AddItem i: Next i

'default
ufb = False
combo_month.Value = getMonth(Month(Date))
combo_year.Value = Year(Date)
ufb = True

Call fillXdates: Call fillCurrentDates: Call formatXs: Call formatAllLabels: Call formatDayinActivate

End Sub

Private Sub UserForm_Terminate()
Application.ScreenUpdating = True
End Sub



Private Sub fillCurrentDates()

Dim date_ As Date
date_ = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), Day(Date))
If Year(date_) = Year(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), 1)) And _
   Month(date_) = Month(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), 1)) Then
Else
  date_ = DateSerial(combo_year.Value, getMonthNum(combo_month.Value) + 1, 0)
End If

Dim i, j, cmwof, d As Byte
cmwof = Weekday(DateSerial(Year(date_), Month(date_), 1), vbMonday) 'cari ayin ilk gununun heftesinin nomresi
cml = Day(DateSerial(Year(date_), Month(date_) + 1, 0))
d = 1 'gunler

'main

'1-ci hefte
For i = cmwof To 7
  Controls("lb_1" & i).Caption = d
  Controls("lb_1" & i).Font.Bold = True
  d = d + 1
Next i

'diger hefteler
For j = 2 To 6
  For i = 1 To 7
    If cml + 1 = d Then GoTo cix
    Controls("lb_" & j & i).Caption = d
    Controls("lb_" & j & i).Font.Bold = True
    d = d + 1
  Next i
Next j
cix:

End Sub

Private Sub fillXdates()

Dim i, j As Byte

For i = 1 To 6
  For j = 1 To 7
    Controls("lb_" & i & j).Caption = "x"
    Controls("lb_" & i & j).FontBold = False
    Controls("lb_" & i & j).Enabled = True
  Next j
Next i

End Sub

Private Sub formatXs()

Dim date_ As Date
date_ = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), Day(Date))
If Year(date_) = Year(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), 1)) And _
   Month(date_) = Month(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), 1)) Then
Else
  date_ = DateSerial(combo_year.Value, getMonthNum(combo_month.Value) + 1, 0)
End If

Dim i, j, d, dd As Integer
Dim pm, py, ldpm As Integer
Dim nm, ny, ndpm As Integer
dd = 1

nm = Month(date_) + 1
If pm = 13 Then ny = Year(date_) + 1: nm = 1 Else ny = Year(date_)
ndpm = Day(DateSerial(ny, nm + 1, 0))

pm = Month(date_) - 1
If pm = 0 Then py = Year(date_) - 1: pm = 12 Else py = Year(date_)
ldpm = Day(DateSerial(py, pm + 1, 0))

'main

'ilk hefte
For i = 1 To 7 'gun
  If Controls("lb_1" & i).Caption = "x" Then
    Controls("lb_1" & i).Enabled = False
    
    For j = ldpm To 1 Step -1
      If Weekday(DateSerial(py, pm, j), vbMonday) = i Then
        Controls("lb_1" & i).Caption = j
        Exit For
      End If
    Next j
    
  End If
Next i

'son 2 hefte
For i = 5 To 6 'hefte
  For j = 1 To 7 'gun
    If Controls("lb_" & i & j).Caption = "x" Then
      Controls("lb_" & i & j).Enabled = False

      For d = 1 To ndpm
        If Weekday(DateSerial(ny, nm, d), vbMonday) = j Then
          Controls("lb_" & i & j).Caption = dd
          dd = dd + 1
          Exit For
        End If
      Next d

    End If
  Next j
Next i

End Sub


















'Label den secim edildikde---------------------------------------------------------------------------------------------------

Private Sub lb_11_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_11.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_11.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_11.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_12_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_12.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_12.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_12.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_13_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_13.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_13.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_13.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_14_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_14.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_14.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_14.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_15_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_15.Caption): Exit Sub
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_15.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_15.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_15.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_16_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_16.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_16.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_16.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_17_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_17.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_17.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_17.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_21_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_21.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_21.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_21.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_22_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_22.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_22.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_22.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_23_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_23.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_23.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_23.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_24_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_24.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_24.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_24.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_25_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_25.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_25.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_25.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_26_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_26.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_26.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_26.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_27_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_27.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_27.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_27.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_31_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_31.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_31.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_31.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_32_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_32.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_32.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_32.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_33_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_33.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_33.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_33.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_34_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_34.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_34.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_34.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_35_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_35.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_35.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_35.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_36_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_36.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_36.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_36.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_37_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_37.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_37.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_37.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_41_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_41.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_41.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_41.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_42_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_42.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_42.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_42.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_43_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_43.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_43.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_43.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_44_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_44.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_44.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_44.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_45_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_45.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_45.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_45.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_46_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_46.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_46.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_46.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_47_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_47.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_47.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_47.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_51_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_51.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_51.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_51.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_52_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_52.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_52.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_52.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_53_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_53.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_53.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_53.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_54_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_54.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_54.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_54.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_55_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_55.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_55.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_55.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_56_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_56.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_56.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_56.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_57_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_57.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_57.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_57.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_61_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_61.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_61.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_61.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_62_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_62.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_62.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_62.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_63_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_63.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_63.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_63.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_64_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_64.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_64.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_64.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_65_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_65.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_65.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_65.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_66_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_66.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_66.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_66.Caption)), "dd.mm.yyyy")
Unload Me
End Sub

Private Sub lb_67_Click()
If DatePickerModule.dpt = "" Then: MsgBox DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_67.Caption): Exit Sub
If DatePickerModule.dpt = 1 Then: Workbooks(DatePickerModule.wrk).Sheets(DatePickerModule.ws).Cells(DatePickerModule.rngR, DatePickerModule.rngC).Value = DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_67.Caption)
If DatePickerModule.dpt = 2 Then: DatePickerModule.uf.Controls(DatePickerModule.c.Name).Value = Format(CDate(DateSerial(combo_year.Value, getMonthNum(combo_month.Value), lb_67.Caption)), "dd.mm.yyyy")
Unload Me
End Sub










'Label uzerine geldikde---------------------------------------------------------------------------------------------------
Private Sub lb_11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(1, 1): End Sub
Private Sub lb_12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(1, 2): End Sub
Private Sub lb_13_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(1, 3): End Sub
Private Sub lb_14_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(1, 4): End Sub
Private Sub lb_15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(1, 5): End Sub
Private Sub lb_16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(1, 6): End Sub
Private Sub lb_17_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(1, 7): End Sub
Private Sub lb_21_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(2, 1): End Sub
Private Sub lb_22_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(2, 2): End Sub
Private Sub lb_23_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(2, 3): End Sub
Private Sub lb_24_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(2, 4): End Sub
Private Sub lb_25_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(2, 5): End Sub
Private Sub lb_26_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(2, 6): End Sub
Private Sub lb_27_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(2, 7): End Sub
Private Sub lb_31_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(3, 1): End Sub
Private Sub lb_32_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(3, 2): End Sub
Private Sub lb_33_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(3, 3): End Sub
Private Sub lb_34_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(3, 4): End Sub
Private Sub lb_35_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(3, 5): End Sub
Private Sub lb_36_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(3, 6): End Sub
Private Sub lb_37_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(3, 7): End Sub
Private Sub lb_41_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(4, 1): End Sub
Private Sub lb_42_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(4, 2): End Sub
Private Sub lb_43_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(4, 3): End Sub
Private Sub lb_44_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(4, 4): End Sub
Private Sub lb_45_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(4, 5): End Sub
Private Sub lb_46_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(4, 6): End Sub
Private Sub lb_47_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(4, 7): End Sub
Private Sub lb_51_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(5, 1): End Sub
Private Sub lb_52_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(5, 2): End Sub
Private Sub lb_53_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(5, 3): End Sub
Private Sub lb_54_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(5, 4): End Sub
Private Sub lb_55_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(5, 5): End Sub
Private Sub lb_56_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(5, 6): End Sub
Private Sub lb_57_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(5, 7): End Sub
Private Sub lb_61_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(6, 1): End Sub
Private Sub lb_62_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(6, 2): End Sub
Private Sub lb_63_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(6, 3): End Sub
Private Sub lb_64_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(6, 4): End Sub
Private Sub lb_65_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(6, 5): End Sub
Private Sub lb_66_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(6, 6): End Sub
Private Sub lb_67_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): Call formatAllLabels: Call mouseOverLabels(6, 7): End Sub


Private Sub formatAllLabels()
'default

Dim i, j As Byte

For i = 1 To 6
  For j = 1 To 7
    Controls("lb_" & i & j).ForeColor = &H80000012
    Controls("lb_" & i & j).BackColor = &H8000000F
  Next j
Next i

End Sub

Private Sub mouseOverLabels(ByVal i As Byte, ByVal j As Byte)
'on mouse lable

Controls("lb_" & i & j).ForeColor = vbWhite
Controls("lb_" & i & j).BackColor = vbBlue
End Sub

Private Sub formatDayinActivate()

Dim i, j As Byte
Dim d As Date
d = Date

For i = 1 To 6
  For j = 1 To 7
    If Controls("lb_" & i & j).Enabled = True And _
       CInt(Controls("lb_" & i & j).Caption) = Day(d) Then
      Controls("lb_" & i & j).ForeColor = vbWhite
      Controls("lb_" & i & j).BackColor = vbBlue
      Exit For
    End If
  Next j
Next i

End Sub



'Diger-----------------------------------------------------------------------------
Private Function getMonthNum(ByVal ay As String) As Byte

If ay = "January" Then getMonthNum = 1
If ay = "February" Then getMonthNum = 2
If ay = "March" Then getMonthNum = 3
If ay = "April" Then getMonthNum = 4
If ay = "May" Then getMonthNum = 5
If ay = "June" Then getMonthNum = 6
If ay = "July" Then getMonthNum = 7
If ay = "August" Then getMonthNum = 8
If ay = "September" Then getMonthNum = 9
If ay = "October" Then getMonthNum = 10
If ay = "November" Then getMonthNum = 11
If ay = "December" Then getMonthNum = 12

End Function

Private Function getMonth(ByVal n As Byte) As String

If n = 1 Then getMonth = "January"
If n = 2 Then getMonth = "February"
If n = 3 Then getMonth = "March"
If n = 4 Then getMonth = "April"
If n = 5 Then getMonth = "May"
If n = 6 Then getMonth = "June"
If n = 7 Then getMonth = "July"
If n = 8 Then getMonth = "August"
If n = 9 Then getMonth = "September"
If n = 10 Then getMonth = "October"
If n = 11 Then getMonth = "November"
If n = 12 Then getMonth = "December"

End Function

Private Sub lb_today_()

If getMonthNum(combo_month.Value) = Month(Date) And _
   CInt(combo_year.Value) = Year(Date) Then
  
  lb_today.Caption = ""
  Call formatDayinActivate
Else
  lb_today.Caption = Format(DateSerial(Year(Date), Month(Date), Day(Date)), "dd.mm.yyyy")
End If

End Sub
