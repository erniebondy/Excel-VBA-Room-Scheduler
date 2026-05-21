VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Room"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pId As Long
Private pName As String
'Private pSchedules() As Schedule
Private pSchedules As Collection

Sub Make(Id As Long, Name As String)
    pId = Id
    pName = Name
    Call SetSchedulesObjects
End Sub

Private Sub SetSchedulesObjects()
    Dim Conn As Connection
    Set Conn = New Connection
    Conn.ConnectionString = ConnectionString
    Conn.Open
    
    Dim Command As Command
    Set Command = New Command
    Command.ActiveConnection = Conn
    Command.CommandText = "SELECT room_id, date, ISNULL(description, '') AS description" & vbNewLine & _
                          "FROM SCHEDULES WHERE room_id = ? ORDER BY date ASC"
    Command.Parameters.Append Command.CreateParameter("P1", adInteger, , , pId)
    
    Dim Rs As Recordset
    Set Rs = Command.Execute
    
    If Rs.EOF Then
        Conn.Close
        'ReDim pSchedules(-1 To -1)
        Set pSchedules = New Collection
        Exit Sub
    End If
    
    Dim DbRows()
    DbRows = Rs.GetRows
    Conn.Close
    
    Dim Rows()
    Rows = TransposeDBArray(DbRows)
    
'    ReDim pSchedules(UBound(Rows))
    
    Dim Row As Long
    For Row = LBound(Rows) To UBound(Rows)
        pSchedules.Add NewSchedule(CLng(Rows(Row, 0)), CDate(Rows(Row, 1)), CStr(Rows(Row, 2)))
        'Set pSchedules(Row) = NewSchedule(CLng(Rows(Row, 0)), CDate(Rows(Row, 1)), CStr(Rows(Row, 2)))
    Next
End Sub

Property Set Schedules(RHS As Collection)
    Set pSchedules = RHS
End Property

Property Get Schedules() As Collection
    Set Schedules = pSchedules
End Property

'Sub SetSchedules(Schedules() As Schedule)
'    pSchedules = Schedules
'End Sub
'
''Property Let Schedules(RHS() As Schedule)
''    pSchedules = RHS
''End Property
'
'Property Get SchedulesItem(Idx As Long) As Schedule
'    Set SchedulesItem = pSchedules(Idx)
'End Property
'
'Property Get Schedules() As Schedule()
'    Schedules = pSchedules
'End Property

Property Let Id(RHS As Long)
    pId = RHS
End Property

Property Get Id() As Long
    Id = pId
End Property

Property Let Name(RHS As String)
    pName = RHS
End Property

Property Get Name() As String
    Name = pName
End Property

Private Sub Class_Initialize()
    Set pSchedules = New Collection
End Sub
