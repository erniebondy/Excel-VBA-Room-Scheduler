Attribute VB_Name = "ModClassFactory"
Option Explicit

Function NewRoom(Id As Long, Name As String) As Room
    Dim Obj As Room
    Set Obj = New Room
    Obj.Make Id, Name
    Set NewRoom = Obj ' Return
End Function

Function NewSchedule(RoomId As Long, DateValue As Date, Description As String)
    Dim Obj As Schedule
    Set Obj = New Schedule
    Obj.Make RoomId, DateValue, Description
    Set NewSchedule = Obj ' Return
End Function
