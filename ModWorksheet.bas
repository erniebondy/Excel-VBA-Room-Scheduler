Attribute VB_Name = "ModWorksheet"
Option Explicit

Sub ClearFormat(Ws As Worksheet, Column As Long)

    Dim Rng As Range
    Set Rng = Ws.Columns(Column)
    
    Rng.ClearFormats
    
End Sub
