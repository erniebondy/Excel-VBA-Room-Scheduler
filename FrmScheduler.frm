VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmScheduler 
   Caption         =   "Meeting Room Scheduler"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11910
   OleObjectBlob   =   "FrmScheduler.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmScheduler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pRefDate As Date
Private WithEvents pWrappers As ControlWrappers

Private pActiveControl As ControlWrapper

Private pSelectedRoom As Room

Property Get SelectedRoom() As Room
    Set SelectedRoom = pSelectedRoom
End Property

Private Sub BtnNextMonth_Click()
    pRefDate = DateAdd("m", 1, pRefDate)
    Call SetMonthLabel(pRefDate)
    Call DrawDayLabels(pRefDate)
End Sub

Private Sub BtnPreviousMonth_Click()
    pRefDate = DateAdd("m", -1, pRefDate)
    Call SetMonthLabel(pRefDate)
    Call DrawDayLabels(pRefDate)
End Sub

Private Sub BtnRoomAdd_Click()
    FrmRoomsConfig.Show
End Sub

Private Sub LboRooms_Click()
    Call HandleLboRooms
End Sub

Private Sub LboRooms_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call HandleLboRooms
End Sub

Private Sub HandleLboRooms()
    Call DrawDayLabels(pRefDate)

    Dim Indexes() As Long
    Indexes = GetSelectedIndexes(LboRooms)

    If UBound(Indexes) < 0 Then
        Set pSelectedRoom = Nothing
        Exit Sub
    End If

    Dim Idx As Long
    Idx = Indexes(0)
    
    Set pSelectedRoom = NewRoom(LboRooms.List(Idx, 0), LboRooms.List(Idx, 1))

    Call SetCalendar(pSelectedRoom, pRefDate)
End Sub

Private Sub mWrappers_ActiveControlClick()
    ' Show menu
    
End Sub

Private Sub mWrappers_ActiveControlMouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    If pActiveControl Is Nothing Then
        Set pActiveControl = pWrappers.ActiveControl
    End If
    
    If Not pActiveControl Is pWrappers.ActiveControl Then
        ' Switching Controls
        pActiveControl.ResetBackColor
        Set pActiveControl = pWrappers.ActiveControl
    End If
    
    If Not pActiveControl.MyLabel Is Nothing Then
        pActiveControl.BackColor = vbRed
    End If
    
    'Debug.Print pWrappers.ActiveControl.MyControl.Name & " MOUSE MOVING"
End Sub

Private Sub UserForm_Initialize()
    
    pRefDate = Date
    Call SetMonthLabel
    Call DrawDayLabels
    Call LoadRoomsLbo(LboRooms)

'    Set mWrappers = New ControlWrappers
'    Set mActiveControl = Nothing
'
'    Dim Wrapper As ControlWrapper
'    Set Wrapper = New ControlWrapper
'    Set Wrapper.MyListBox = LboRooms
'    mWrappers.Add Wrapper
'
'    Set Wrapper = New ControlWrapper
'    Set Wrapper.MyLabel = LblFile
'    mWrappers.Add Wrapper
'
'    Set Wrapper = New ControlWrapper
'    Set Wrapper.MyForm = Me
'    mWrappers.Add Wrapper

    
End Sub
