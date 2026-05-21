VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Schedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pRoomId As Long
Private pDateValue As Date
Private pDescription As String

Sub Make(RoomId As Long, DateValue As Date, Description As String)
    pRoomId = RoomId
    pDateValue = DateValue
    pDescription = Description
End Sub

Public Property Let RoomId(RHS As Long)
    pRoomId = RHS
End Property

Public Property Get RoomId() As Long
    RoomId = pRoomId
End Property

Public Property Get DateValue() As Date
    DateValue = pDateValue
End Property

Public Property Get Description() As String
    Description = pDescription
End Property
