Attribute VB_Name = "ESToolKit"
Public TableCache As Dictionary
Public Sub TestMe()
    cacheMe = "aaa"
End Sub

Public Sub SetF(fieldName As String, val)
    Dim Address As String
    Address = GetFAddr(fieldName)
    If (InStr(Address, ".") > 0) Then
        Dim arr
        arr = Split(Address, ".")
        Sheets(arr(0)).Range(arr(1)).Value = val
    Else
        ActiveSheet.Range(Address).Value = val
    End If
End Sub

Public Function GetF(fieldName As String)
    Dim Address As String
    Address = GetFAddr(fieldName)
    Dim firstCell As String
    If (InStr(Address, ".") > 0) Then
        Dim arr
        arr = Split(Address, ".")
        firstCell = getFirstCellAddress(CStr(arr(1)))
        GetF = Sheets(arr(0)).Range(firstCell).Value
    Else
        firstCell = getFirstCellAddress(Address)
        GetF = ActiveSheet.Range(firstCell).Value
    End If
End Function

Public Function GetFAddr(fieldName As String)
    Dim obj As Object
    Set obj = Application.COMAddIns("ESClient10.Connect").Object
    Dim Address As String, startRow As Long, startCol As Long, EndRow As Long, EndCol As Long
    obj.GetFieldAddress fieldName, Address, startRow, startCol, EndRow, EndCol
    Set obj = Nothing
    GetFAddr = Address
End Function

Public Function GetFRange(fieldName As String) As Range
    Dim Address As String
    Address = GetFAddr(fieldName)
    Dim firstCell As String
    If (InStr(Address, ".") > 0) Then
        Dim arr
        arr = Split(Address, ".")
        Set GetFRange = Sheets(arr(0)).Range(arr(1))
    Else
        Set GetFRange = ActiveSheet.Range(Address)
    End If
End Function

Public Sub Focus(fieldName As String)
    Dim Address As String
    Address = GetFAddr(fieldName)
    Dim firstCell As String
    If (InStr(Address, ".") > 0) Then
        Dim arr
        arr = Split(Address, ".")
        firstCell = getFirstCellAddress(CStr(arr(1)))
        Sheets(arr(0)).Activate
        Sheets(arr(0)).Range(firstCell).Select
    Else
        firstCell = getFirstCellAddress(Address)
        ActiveSheet.Range(firstCell).Select
    End If
    
End Sub

Public Function GetF_Dt(fieldName As String, rowIndex As Long)
    Dim Address As String
    Address = GetFAddr(fieldName)

    Dim firstCell As String
    If (InStr(Address, ".") > 0) Then
        Dim arr
        arr = Split(Address, ".")
        If (Sheets(arr(0)).Range(arr(1)).Rows.Count < rowIndex) Then
            MsgBox "Exec GetF_Dt Error,over max row"
            Exit Function
        End If
        GetF_Dt = Sheets(arr(0)).Range(arr(1)).Cells(rowIndex, 1).Value
    Else
        If (ActiveSheet.Range(Address).Rows.Count < rowIndex) Then
            MsgBox "Exec GetF_Dt Error,over max row"
            Exit Function
        End If
        GetF_Dt = ActiveSheet.Range(Address).Cells(rowIndex, 1).Value
    End If
End Function

Public Sub SetF_Dt(fieldName As String, rowIndex As Long, val)
    Dim Address As String
    Address = GetFAddr(fieldName)
    If (InStr(Address, ".") > 0) Then
        Dim arr
        arr = Split(Address, ".")

        If (Sheets(arr(0)).Range(arr(1)).Rows.Count < rowIndex) Then
            MsgBox "Exec SetF_Dt Error,over max row"
            Exit Sub
        End If

        Sheets(arr(0)).Range(arr(1)).Cells(rowIndex, 1).Value = val
    Else
        If (ActiveSheet.Range(Address).Rows.Count < rowIndex) Then
            MsgBox "Exec SetF_Dt Error,over max row"
            Exit Sub
        End If
        ActiveSheet.Range(Address).Cells(rowIndex, 1).Value = val
    End If
End Sub

Public Function GetRelateLine(absLine As Long, tableName As String) As Long

    Dim startRow As Long, EndRow As Long, shtIndex As Integer
    Dim rng As Range
    Set rng = GetTableRange(tableName, shtIndex)
    startRow = getFirstRowByRng(rng)
    EndRow = getLastRowByRng(rng)

    If (absLine <= EndRow) Then
        GetRelateLine = absLine - startRow + 1
    Else
        MsgBox "Exec GetRelateLine Error,over max row"
        Exit Function
    End If
End Function
Public Function GetAbsLine(relLine As Long, tableName As String) As Long

    Dim startRow As Long, EndRow As Long, shtIndex As Integer
    Dim rng As Range
    Set rng = GetTableRange(tableName, shtIndex)
    startRow = getFirstRowByRng(rng)
    EndRow = getLastRowByRng(rng)

    If (relLine <= EndRow - startRow + 1) Then
        GetAbsLine = relLine + startRow - 1
    Else
        MsgBox "Exec GetAbsLine Error,over max row"
        Exit Function
    End If
End Function



Public Function GetTableRange(tableName As String, ByRef shtIndex As Integer) As Range
    If (TableCache Is Nothing) Then
        Set TableCache = New Dictionary
    End If
    Dim RngNo As Integer, si As Integer
    If (TableCache.Exists(tableName)) Then
        si = TableCache(tableName)(1)
        RngNo = TableCache(tableName)(2)
    Else
        Dim rs As ADODB.Recordset
        Dim obj As Object
        Set obj = Application.COMAddIns("ESClient10.Connect").Object
        Dim errMsg As String
        If (obj.ExecQryProc("p_ESTK_FilterInfo", rs, errMsg, tableName) = True) Then
            Set obj = Nothing
            If (rs.BOF And rs.EOF) Then
                Set GetTableRange = Nothing
                Exit Function
            Else
                rs.MoveFirst
                si = rs("SheetId")
                RngNo = rs("RttId")
                Dim t(1 To 2) As Integer
                t(1) = si
                t(2) = RngNo
                
                TableCache.Add tableName, t

            End If
        Else
            Set obj = Nothing
            Set GetTableRange = Nothing
            Exit Function
        End If
    End If
    shtIndex = si
    Set GetTableRange = Sheets(si).Range("_EST" + CStr(RngNo))
    
End Function
Public Sub AddRows(tableName As String, rowCount As Long)
    Dim rng As Range, shtIndex As Integer
    Set rng = GetTableRange(tableName, shtIndex)
    Dim obj As Object
    Set obj = Application.COMAddIns("ESClient10.Connect").Object
    obj.insertRow shtIndex, getLastRowByRng(rng), rowCount
    Set obj = Nothing
End Sub
Public Sub ClearRows(tableName As String)
    Dim obj As Object
    Set obj = Application.COMAddIns("ESClient10.Connect").Object
    Dim shtIndex As Long, rng As Range
    Set rng = GetTableRange(tableName, shtIndex)
    obj.deleteRow shtIndex, getFirstRowByRng(rng), getRowCount(tableName)
    Set obj = Nothing
End Sub
Public Sub DelOneRow(tableName As String, recordIndex As Long)

    Dim rng As Range, shtIndex As Integer
    Set rng = GetTableRange(tableName, shtIndex)

    Dim obj As Object
    Set obj = Application.COMAddIns("ESClient10.Connect").Object
    obj.deleteRow shtIndex, GetAbsLine(recordIndex, tableName), 1
    Set obj = Nothing
End Sub
Public Sub DelRowsByFilter(tableName As String, filterDict As Dictionary)
    Dim rng As Range, shtIndex As Integer
    Set rng = GetTableRange(tableName, shtIndex)
    Dim i As Long, firstRow As Long, lastRow As Long
    firstRow = getFirstRowByRng(rng)
    lastRow = getLastRowByRng(rng)
    Dim k As Long
    k = 1
    Dim delList As Collection
    Set delList = New Collection
    For i = firstRow To lastRow
        If (IsPass(k, filterDict)) Then
            delList.Add k
        End If
        k = k + 1
    Next
    If (delList.Count > 0) Then
        For i = delList.Count To 1 Step -1
            DelOneRow tableName, delList(i)
        Next
    End If
End Sub
Private Function IsPass(dataLine As Long, filterDict As Dictionary) As Boolean
    If (filterDict.Count = 0) Then
        IsPass = True
        Exit Function
    End If
    Dim i As Integer
    Dim key As String, val
    For i = 0 To filterDict.Count - 1
        key = filterDict.keys(i)
        val = filterDict(key)
        If (GetF_Dt(key, dataLine) <> val) Then
            IsPass = False
            Exit Function
        End If
    Next
    IsPass = True
End Function
'Public Sub AddRowByCount(sht As Worksheet, rngName As String, rowCount As Long, fixCount As Long)
'    Dim obj As Object
'    Set obj = Application.COMAddIns("ESClient10.Connect").Object
'    If (rowCount > fixCount) Then
'            obj.insertRow sht.Index, getLastRow(sht, rngName), rowCount - fixCount
'    End If
'    Set obj = Nothing
'End Sub
Public Function getRowCountByRng(rng As Range) As Long
    getRowCountByRng = getLastRowByRng(rng) - getFirstRowByRng(rng)
End Function
Public Function getRowCount(tableName As String) As Long
    Dim rng As Range
    Set rng = GetTableRange(tableName)
    getRowCount = getRowCountByRng(rng)
End Function
'Public Function getRowCount(fieldName As String) As Long
'    Dim obj As Object
'    Set obj = Application.COMAddIns("ESClient10.Connect").Object
'    Dim Address As String, startRow As Long, startCol As Long, EndRow As Long, EndCol As Long
'    obj.GetFieldAddress fieldName, Address, startRow, startCol, EndRow, EndCol
'    getRowCount = EndRow - startRow + 1
'End Function


Private Function getFirstCellAddress(addr As String)
    Dim arr
    arr = Split(addr, ":")
    getFirstCellAddress = arr(0)
End Function

Public Function getLastRow(tableName As String) As Long
    Dim rng As Range, shtIndex As Integer
    Set rng = GetTableRange(tableName, shtIndex)
    getLastRow = getLastRowByRng(rng)
End Function
Public Function getLastRowByRng(rng As Range) As Long
    Dim lCount As Long
    lCount = rng.Cells.Count
    getLastRowByRng = rng.Cells(lCount).Row
End Function

Public Function getFirstRow(tableName As String) As Long
    Dim rng As Range, shtIndex As Integer
    Set rng = GetTableRange(tableName, shtIndex)
    getFirstRow = getFirstRowByRng(rng)
End Function
Public Function getFirstRowByRng(rng As Range) As Long
    getFirstRowByRng = rng.Cells(1).Row
End Function
'Public Function getLastRow(sht As Worksheet, rngName As String) As Long
'    Dim lCount As Long
'    lCount = sht.Range(rngName).Cells.Count
'    getLastRow = sht.Range(rngName).Cells(lCount).Row
'End Function
'Public Function getFirstRow(sht As Worksheet, rngName As String) As Long
'    Dim lCount As Long
'    lCount = sht.Range(rngName).Cells.Count
'    getFirstRow = sht.Range(rngName).Cells(1).Row
'End Function


Public Function IsDesign() As Boolean
    If (InStr(Application.Caption, "设计：") = 0 Or InStr(Application.Caption, "设计:") = 0) Then
        IsDesign = False
    Else
        IsDesign = True
    End If
End Function

'对一个集合dictionary进行排序(从小到大),排序的依据是Value,排序后的集合返回
Public Sub SortDictionaryDesc(dict As Dictionary)
    Dim newdict As New Dictionary
    Dim minKey As String
    While (dict.Count > 0)
        minKey = GetMinFromDict(dict)
        newdict.Add minKey, dict(minKey)
        dict.Remove minKey
    Wend
    Set dict = newdict
End Sub

'得到dictionary中最小Value的key值
'key值为str类型
Public Function GetMinFromDict(dict As Dictionary) As String
    Dim i As Integer
    Dim minVal As Single
    minVal = 9999
    Dim recKey As String
    recKey = ""
    For i = 0 To dict.Count - 1
        If (minVal > dict(dict.keys(i))) Then
            minVal = dict(dict.keys(i))
            recKey = dict.keys(i)
        End If
    Next
    GetMinFromDict = recKey
End Function

