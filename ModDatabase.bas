Attribute VB_Name = "ModDatabase"
Option Explicit

'Const DATABASE_WORKBOOK_PATH As String = "C:\Users\ernie\Documents\Programming\VBA"
'Const DATABASE_WORKBOOK_NAME As String = "MRS-DB.xlsx"

Public Property Get ConnectionString() As String
    ConnectionString = "Provider=MSOLEDBSQL;Server=localhost\SQLEXPRESS;Database=RoomScheduler;UID=excel-vba;PWD=123;"
    'ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                       "Data Source=" & DATABASE_WORKBOOK_PATH & Application.PathSeparator & DATABASE_WORKBOOK_NAME & ";" & _
                       "Extended Properties=""Excel 12.0;HDR=YES"";"
    'ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                       "Data Source=" & DATABASE_WORKBOOK_PATH & Application.PathSeparator & DATABASE_WORKBOOK_NAME & ";" & _
                       "Extended Properties=""Excel 12.0 Macro;HDR=YES;IMEX=0"";"
End Property

Function GetNextId(Conn As Connection, TableName As String) As Long
    Dim LastId As Long
    LastId = GetLastId(Conn, TableName)
    GetNextId = LastId + 1
End Function

Function GetLastId(Conn As Connection, TableName As String) As Long
    Dim Command As Command
    Set Command = New Command
    Command.ActiveConnection = Conn
    Command.CommandText = "SELECT TOP 1 id FROM [" & TableName & "$] ORDER BY id DESC"
    
    Dim Rs As Recordset
    Set Rs = Command.Execute
    
    Dim LastId As Long
    LastId = Rs!Id
    
    Set Rs = Nothing
    Set Command = Nothing
    
    GetLastId = LastId ' Return
End Function

Sub BangAccessor()

    Dim Coll As New Collection
    Coll.Add "wordA", "A"
    Coll.Add "wordB", "B"
    
    Dim Dic As New Dictionary
    Dic.Add "A", "wordAA"
    Dic.Add "B", "wordBB"
    
    Debug.Print "|"; Coll!A; "|"    ' Output: |wordA|
    Debug.Print "|"; Dic!A; "|"     ' Output: |wordAA|
    Debug.Print "|"; Coll!C; "|"    ' Output: ERROR
    Debug.Print "|"; Dic!C; "|"     ' Output: ||

End Sub


Sub AA()
    Dim Conn As Connection
    Set Conn = New Connection
    Conn.ConnectionString = ConnectionString
    Conn.Open
    
    Dim Command As Command
    Set Command = New Command
    Command.ActiveConnection = Conn
    Command.CommandText = "SELECT * FROM ROOMS"
    
    Dim Rs As Recordset
    'Set Rs = Conn.OpenSchema(adSchemaTables)
    Set Rs = Command.Execute
    
    Do While Not Rs.EOF
        Dim Col As Long
        For Col = 0 To Rs.Fields.Count - 1
            Debug.Print Rs(Col).Name & ": '" & Rs(Col).Value & "'"
        Next
        Rs.MoveNext
    Loop
        
    Conn.Close
End Sub
