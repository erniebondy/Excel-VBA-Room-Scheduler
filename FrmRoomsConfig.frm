VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmRoomsConfig 
   Caption         =   "Rooms"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5910
   OleObjectBlob   =   "FrmRoomsConfig.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmRoomsConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum RoomsColumn
    enId
    enName
End Enum

Private Sub BtnRoomAdd_Click()

    Dim RoomName As String
    RoomName = Trim(TxtRoomName.Value)
    
    If RoomName = vbNullString Then
        MsgBox "No room name!", vbCritical
        Exit Sub
    End If
    
    Dim Ok As Boolean
    Ok = RoomAdd(RoomName)
    
    If Ok Then
        TxtRoomName.Value = vbNullString
        Call FormLoad
        Exit Sub
    End If
    
    MsgBox "Room NOT added!", vbCritical
End Sub

Private Sub BtnRoomDelete_Click()
    Dim Indexes() As Long
    Indexes = GetSelectedIndexes(LboRooms)
    
    If UBound(Indexes) < 0 Then
        MsgBox "No room selected!", vbCritical
        Exit Sub
    End If
    
    Dim Idx As Long
    Idx = Indexes(0)
    
    Dim Id As Long
    Id = LboRooms.List(Idx, RoomsColumn.enId)
    
    Dim Ok As Boolean
    Ok = RoomDelete(Id)
    
End Sub

Private Sub BtnRoomEdit_Click()
    Dim NewRoomName As String
    NewRoomName = Trim(TxtRoomName.Value)
    
    If NewRoomName = vbNullString Then
        MsgBox "No room name!", vbCritical
        Exit Sub
    End If
    
    Dim SelectedRoom As Room
    Set SelectedRoom = Nothing
    
    Dim Idx As Long
    For Idx = 0 To LboRooms.ListCount - 1
        If LboRooms.Selected(Idx) Then
            Set SelectedRoom = NewRoom(LboRooms.List(Idx, RoomsColumn.enId), vbNullString)
            Exit For
        End If
    Next
    
    If SelectedRoom Is Nothing Then
        MsgBox "No room selected!", vbCritical
        Exit Sub
    End If
    
    SelectedRoom.Name = NewRoomName
    
    Dim Ok As Boolean
    Ok = RoomEdit(SelectedRoom)

    If Ok Then
        TxtRoomName.Value = vbNullString
        Call FormLoad
        Exit Sub
    End If
    
    MsgBox "Room NOT updated!", vbCritical
End Sub

Private Sub UserForm_Initialize()
    
'    On Error Resume Next
'    Debug.Print "HInstanncePtr "; Application.HinstancePtr
'    Debug.Print "Hwnd "; Application.Hwnd
'    On Error GoTo 0
    
    Call FormLoad
End Sub

Private Sub FormLoad()
    Call LoadRoomsLbo(LboRooms)
End Sub
