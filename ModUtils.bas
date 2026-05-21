Attribute VB_Name = "ModUtils"
Option Explicit

Function TransposeDBArray(SrcArr()) As Variant()
    '--------'--------'
    ' SrcArr ' NewArr '
    '--------'--------'
    ' C1R1   ' R1C1   '
    ' C1R2   ' R1C2   '
    ' C1R3   ' R2C1   '
    ' C2R1   ' R2C2   '
    ' C2R2   ' R3C1   '
    ' C2R3   ' R3C2   '
    '--------'--------'
    
    Dim RowCount As Long
    RowCount = UBound(SrcArr, 2) + 1
    
    Dim ColumnCount As Long
    ColumnCount = UBound(SrcArr) + 1
    
    Dim NewArr()
    ReDim NewArr(RowCount - 1, ColumnCount - 1)
    
    Dim Idx As Long
    For Idx = 0 To (RowCount * ColumnCount) - 1
        Dim Row As Long
        Row = Idx \ ColumnCount
        
        Dim Column As Long
        Column = Idx Mod ColumnCount
        
        NewArr(Row, Column) = SrcArr(Column, Row)
    Next
        
    TransposeDBArray = NewArr
End Function

Function GetSelectedIndexes(Lbo As MSForms.ListBox) As Long()
    Dim Indexes() As Long
    ReDim Indexes(Lbo.ListCount)
    
    Dim Idx As Long
    Dim Idx2 As Long
    Idx2 = 0
    For Idx = 0 To Lbo.ListCount - 1
        If Lbo.Selected(Idx) Then
            Indexes(Idx2) = Idx
            Idx2 = Idx2 + 1
        End If
    Next
    
    If Idx2 > 0 Then
        ReDim Preserve Indexes(Idx2 - 1)
    Else
        ReDim Indexes(-1 To -1)
    End If
    
    GetSelectedIndexes = Indexes ' Return
End Function

'Function ForEach(Items, Optional Idx As Long = 0) As Object
'    Static ForEachIdx As Long
'    If Idx = 0 Then
'        ForEachIdx = 0
'    End If
'
'    If ForEachIdx > UBound(Items) Then
'        ForEachIdx = -1
'        Set ForEach = Nothing: Exit Function ' Return
'    End If
'
'    Set ForEach = Items(ForEachIdx)
'    ForEachIdx = ForEachIdx + 1
'End Function
