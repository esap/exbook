Attribute VB_Name = "模块1"

Public Sub DirectCopyRs(sht As Worksheet, rngName As String, rs As ADODB.Recordset, Optional fixRowCount As Integer = 2)
    ClearData sht, rngName, fixRowCount
    AddRowByCount sht, rngName, rs.RecordCount, fixRowCount
    sht.Range(getFirstCellAddr(sht, rngName)).CopyFromRecordset rs
End Sub
'mcol格式-----       B:FieldName  例如： B:ProductID  ;  C:ProductName  ;  D:Qty  ;  E:Price
'如果某列需要填入空白字符，用 FILLBLANK  例如：L:FILLBLANK
'如果某列为公式构成，用 "FOMULA-"，其中rs的字段用@FieldName@  ----未实现
Public Sub MapCopyRs(sht As Worksheet, rngName As String, rs As ADODB.Recordset, mcol As Collection, Optional fixRowCount As Integer = 2)
    ClearData sht, rngName, fixRowCount
    AddRowByCount sht, rngName, rs.RecordCount, fixRowCount
    Dim i As Integer, j As Integer
    Dim fRow As Integer
    j = getFirstRow(sht, rngName)
    Do While rs.EOF = False
        For i = 1 To mcol.Count
            If (ParseField(mcol(i)) = "FILLBLANK") Then
                sht.Range(ParseCol(mcol(i)) + CStr(j)) = ""
            'ElseIf (Left(ParseField(mcol(i)), 7) = "FOMULA-") Then
            '    sht.Range(ParseCol(mcol(i)) + CStr(j)) = ParseFomula(rs, ParseField(mcol(i)))
            Else
                sht.Range(ParseCol(mcol(i)) + CStr(j)) = rs(ParseField(mcol(i)))
            End If
        Next
        j = j + 1
        rs.MoveNext
    Loop
End Sub
Private Function ParseFomula(rs As ADODB.Recordset, strs As String) As Object
        
End Function
Private Function ParseCol(s As String) As String
    ParseCol = Split(s, ":")(0)
End Function
Private Function ParseField(s As String) As String
    ParseField = Split(s, ":")(1)
End Function

Private Sub ClearData(sht As Worksheet, rngName As String, fixCount As Integer)
    Dim obj As Object
    Set obj = Application.COMAddIns("ESClient10.Connect").Object
    Dim shtIndex As Integer
    shtIndex = sht.Index
    lrow = getLastRow(sht, rngName)
    'Remove all rows
    obj.deleteRow shtIndex, getFirstRow(sht, rngName), getRowCount(sht, rngName)
    'Check FixCount when need to Adjust
    If (getRowCount(sht, rngName) < fixCount) Then
        obj.insertRow shtIndex, getFirstRow(sht, rngName), fixCount - getRowCount(sht, rngName)
    End If
    Set obj = Nothing
End Sub

Private Sub AddRowByCount(sht As Worksheet, rngName As String, rowCount As Integer, fixCount As Integer)
    Dim obj As Object
    Set obj = Application.COMAddIns("ESClient10.Connect").Object
    If (rowCount > fixCount) Then
            obj.insertRow sht.Index, getLastRow(sht, rngName), rowCount - fixCount
    End If
    Set obj = Nothing
End Sub
Private Function getLastRow(sht As Worksheet, rngName As String) As Integer
    Dim lCount As Integer
    lCount = sht.Range(rngName).Cells.Count
    getLastRow = sht.Range(rngName).Cells(lCount).Row
End Function
Private Function getFirstRow(sht As Worksheet, rngName As String) As Integer
    Dim lCount As Integer
    lCount = sht.Range(rngName).Cells.Count
    getFirstRow = sht.Range(rngName).Cells(1).Row
End Function
Private Function getFirstCellAddr(sht As Worksheet, rngName As String) As String
    getFirstCellAddr = sht.Range(rngName).Cells(1).Address
End Function
Private Function getLastCol(sht As Worksheet, rngName As String) As Integer
    Dim lCount As Integer
    lCount = sht.Range(rngName).Cells.Count
    getLastCol = sht.Range(rngName).Cells(lCount).Column
End Function
Private Function getRowCount(sht As Worksheet, rngName As String) As Integer
    getRowCount = getLastRow(sht, rngName) - getFirstRow(sht, rngName) + 1
End Function

Public Sub UnitTest1()
    MsgBox Application.Evaluate("34+23")
    
End Sub
