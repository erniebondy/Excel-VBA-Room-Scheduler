Attribute VB_Name = "ModRoomsConfig"
Option Explicit

Function RoomDelete(Id As Long) As Boolean
    Dim Conn As Connection
    Set Conn = New Connection
    Conn.ConnectionString = ModDatabase.ConnectionString
    Conn.Open
    
    Dim Command As Command
    Set Command = New Command
    Command.ActiveConnection = Conn
    
    Command.CommandText = "DELETE FROM ROOMS WHERE id = ?"
    Command.Parameters.Append Command.CreateParameter("Id", adInteger, , , Id)
    
    Dim RecordsAffected As Long
    
    Dim Rs As Recordset
    Set Rs = Command.Execute(RecordsAffected)
    
    Conn.Close
    RoomDelete = (RecordsAffected > 0) ' Return
End Function

Function RoomEdit(NewRoom As Room) As Boolean
    Dim Conn As Connection
    Set Conn = New Connection
    Conn.ConnectionString = ModDatabase.ConnectionString
    Conn.Open
    
    Dim Command As Command
    Set Command = New Command
    Command.ActiveConnection = Conn
    
    Command.CommandText = "UPDATE ROOMS SET name = ? WHERE id = ?"
    Command.Parameters.Append Command.CreateParameter("Name", adVarWChar, , Len(NewRoom.Name), NewRoom.Name)
    Command.Parameters.Append Command.CreateParameter("Id", adInteger, , , NewRoom.Id)
    
    Dim RecordsAffected As Long
    
    Dim Rs As Recordset
    Set Rs = Command.Execute(RecordsAffected)
    
    Conn.Close
    
    Call ClearFormat(ShRooms, 2)
    
    RoomEdit = (RecordsAffected > 0) ' Return
End Function

Function RoomAdd(RoomName As String) As Boolean
    Dim Conn As Connection
    Set Conn = New Connection
    Conn.ConnectionString = ModDatabase.ConnectionString
    Conn.Open
    
    Dim Command As Command
    Set Command = New Command
    Command.ActiveConnection = Conn
    
    Command.CommandText = "SELECT COUNT(id) AS count FROM ROOMS WHERE name = ?"
    Command.Parameters.Append Command.CreateParameter("Name", adVarWChar, , Len(RoomName), RoomName)
    
    Dim Rs As Recordset
    Set Rs = Command.Execute
    
    Dim RowCount As Long
    RowCount = Rs("count").Value ' RowCount = Rs!Count
    
    If RowCount > 0 Then
        MsgBox RoomName & " already exists!", vbCritical
        Conn.Close
        RoomAdd = False: Exit Function
    End If
    
    Set Command = New Command
    Command.ActiveConnection = Conn
    Command.CommandText = "INSERT INTO ROOMS VALUES(?)"
    Command.Parameters.Append Command.CreateParameter("Name", adVarChar, , Len(RoomName), RoomName)
    Command.Execute
    
    Conn.Close
    RoomAdd = True
End Function

Sub LoadRoomsLbo(Lbo As MSForms.ListBox)
    
    Lbo.Clear
    
    Dim Conn As Connection
    Set Conn = New Connection
    Conn.ConnectionString = ConnectionString
    
    Dim Command As Command
    Set Command = New Command
    
    Conn.Open
    Command.ActiveConnection = Conn
    Command.CommandText = "SELECT * FROM ROOMS"
    
    Dim Rs As Recordset
    Set Rs = Command.Execute
    
    Debug.Assert Not Rs.EOF
    
    Dim Rows()
    Rows = Rs.GetRows()
    
    Conn.Close
    
    Lbo.List = TransposeDBArray(Rows)
    
End Sub










