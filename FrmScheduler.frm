VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmScheduler 
   Caption         =   "Meeting Room Scheduler"
   ClientHeight    =   5565
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

Private Sub UserForm_Initialize()
    mRefDate = SetMonthLabel
    Call DrawDayLabels
End Sub


