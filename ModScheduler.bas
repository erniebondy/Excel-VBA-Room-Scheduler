Attribute VB_Name = "ModScheduler"
Option Explicit

'ListBox RowSource FORM_DATA!ROOMS

Const DAY_LABEL_WIDTH As Single = 30
Const DAY_LABEL_HEIGHT As Single = 24

Dim Wrappers As Collection

Sub AAA()

    Dim Frm As Object
    Set Frm = FrmScheduler






End Sub

Function SetMonthLabel(Optional RefDate As Date) As Date
    If RefDate = Empty Then
        RefDate = Date
    End If
    FrmScheduler.LblMonth.Caption = MonthName(Month(RefDate))
    SetMonthLabel = RefDate ' Return
End Function

Sub DrawDayLabels(Optional RefDate As Date)

    Set Wrappers = New Collection

    If RefDate = Empty Then
        RefDate = Date
    End If

    Dim FirstDate As Date
    FirstDate = DateSerial(Year(RefDate), Month(RefDate), 1)
    
    Dim BeginDay As Integer
    BeginDay = Weekday(FirstDate)
    
    Dim EndDate As Date
    EndDate = DateAdd("d", -1, DateAdd("m", 1, FirstDate))
    
    Dim TotalDays As Integer
    TotalDays = Day(EndDate)
    
    Dim Col As Integer
    Col = BeginDay - 1
    
    Dim Row As Integer
    Row = 0
    
    Const DAY_LABEL_TOP_OFFSET As Single = 18
    Const DAY_LABEL_LEFT_OFFSET As Single = 0
    
    FrmScheduler.FrCalendar.Enabled = False
    Do While FrmScheduler.FrCalendar.Controls.Count > 7
        FrmScheduler.FrCalendar.Controls.Remove FrmScheduler.FrCalendar.Controls.Count - 1
    Loop
    FrmScheduler.FrCalendar.Enabled = True
        
    Dim Idx As Integer
    For Idx = 0 To TotalDays - 1
        Dim Lbl As MSForms.Label
        Set Lbl = MakeDayLabel()
        Lbl.Name = "LblDay_" & (Idx + 1)
        Lbl.Top = DAY_LABEL_TOP_OFFSET + (Row * DAY_LABEL_HEIGHT)
        Lbl.Left = DAY_LABEL_LEFT_OFFSET + (Col * DAY_LABEL_WIDTH)
        Lbl.Caption = (Idx + 1)
        
        Dim Wrapper As LabelWrapper
        Set Wrapper = New LabelWrapper
        Set Wrapper.MyLabel = Lbl
        Wrappers.Add Wrapper, Lbl.Name
        
        Col = Col + 1
        If Col Mod 7 = 0 Then
            Col = 0
            Row = Row + 1
        End If
    Next
End Sub

Private Function MakeDayLabel() As MSForms.Label
    Dim Lbl As MSForms.Label
    Set Lbl = FrmScheduler.FrCalendar.Controls.Add("Forms.Label.1", , True)
    Lbl.BorderStyle = fmBorderStyleSingle
    Lbl.TextAlign = fmTextAlignCenter
    Lbl.Width = DAY_LABEL_WIDTH
    Lbl.Height = DAY_LABEL_HEIGHT
    Set MakeDayLabel = Lbl
End Function
