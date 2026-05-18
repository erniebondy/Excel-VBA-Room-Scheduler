Attribute VB_Name = "ModDatabase"
Option Explicit

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
    Conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                            "Data Source=" & ThisWorkbook.FullName & ";" & _
                            "Extended Properties=""Excel 12.0;HDR=YES;IMEX=0"";"
    Conn.Open
    
    Dim Command As Command
    Set Command = New Command
    Command.ActiveConnection = Conn
    Command.CommandText = "SELECT * FROM [ROOMS$]"
    
    Dim Rs As Recordset
    'Set Rs = Conn.OpenSchema(adSchemaTables)
    Set Rs = Command.Execute
    
    Do While Not Rs.EOF
        Dim Col As Long
        For Col = 0 To Rs.Fields.Count - 1
            Debug.Print Rs(Col).Name & ": " & Rs(Col).Value
        Next
        Rs.MoveNext
    Loop
        
    Conn.Close
End Sub
