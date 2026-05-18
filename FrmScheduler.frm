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

Private mRefDate As Date
Private mWrappers As Collection

Private Sub BtnNextMonth_Click()
    mRefDate = DateAdd("m", 1, mRefDate)
    Call SetMonthLabel(mRefDate)
    Call DrawDayLabels(mRefDate)
End Sub

Private Sub BtnPreviousMonth_Click()
    mRefDate = DateAdd("m", -1, mRefDate)
    Call SetMonthLabel(mRefDate)
    Call DrawDayLabels(mRefDate)
End Sub

Private Sub LblFile_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'
End Sub

Private Sub LboRooms_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Button 1 = Left 2 = Right
End Sub

Private Sub UserForm_Initialize()

    Set mWrappers = New Collection

    Dim Wrapper As ControlWrapper
    Set Wrapper = New ControlWrapper
    Set Wrapper.MyListBox = LboRooms
    mWrappers.Add Wrapper
    Set Wrapper = New ControlWrapper
    Set Wrapper.MyLabel = LblFile
    mWrappers.Add Wrapper
    Set Wrapper = New ControlWrapper
    Set Wrapper.MyForm = Me
    mWrappers.Add Wrapper

    mRefDate = Date
    Call SetMonthLabel
    Call DrawDayLabels
End Sub


