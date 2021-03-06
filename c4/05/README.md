# 4.5 VBA的图片接口怎么用？怎么实现点哪儿插哪儿
*注意： 9.4以上才有图片插入VBA接口*

```vb
Sub pic_insert()
	Dim fn
	fn = Application.GetOpenFilename()
	Application.COMAddIns("ESClient10.Connect").Object.AddPicture fn, 1, ActiveCell.Row, ActiveCell.Column
End Sub
```

## 定点插入示例（R5C7）
```vb
Sub IPIC()
'  function AddPicture(path:BSTR; sh:I2; r:I4; c:I4);  ES vba 接口
    Dim fn                         '存放打开的文件
    '弹出文件打开选框
    fn = Application.GetOpenFilename("图片文件(*.JPG;*.PNG;*.BMP),*.JPG;*.PNG;*.BMP", , "打开（可多选）")
    If fn = "" Then Exit Sub                                     '用户未选择文件
    Cells(5, 7).Select
    Application.COMAddIns("ESClient10.Connect").Object.AddPicture fn, 1, 5, 7 ' 插入图片
End Sub
```

## 批量上传示例
![](./4.5.jpg)
```vb
Sub IPIC()
'  function AddPicture(path:BSTR; sh:I2; r:I4; c:I4);  ES vba 接口
    Dim fn, j%                             '文件数组，用于存放打开的文件列表
    Dim oAdd As Object                     'ES 对象
    Set oAdd = Application.COMAddIns("ESClient10.Connect").Object
    '弹出文件打开选框，用户可多选
    fn = Application.GetOpenFilename("图片文件(*.JPG;*.PNG;*.BMP),*.JPG;*.PNG;*.BMP", , "打开（可多选）", , True)
    If IsArray(fn) Then                                      '用户选择了文件时，开始执行复制转换操作
        For j = 1 To UBound(fn)                               '遍历每个文件
            Cells(2 + j, 2) = Left(Dir(fn(j)), InStrRev(Dir(fn(j)), ".") - 1)
            Cells(2 + j, 3).Select
            oAdd.AddPicture fn(j), 1, 2 + j, 3  ' 插入图片
        Next j
'        oAdd.saveCase , True, False             '保存，激活回写
    End If
    Set oAdd = Nothing
End Sub
Sub pic_clear()
    [a1].Select
    Application.COMAddIns("ESClient10.Connect").Object.execQuery "Clear"
End Sub
Sub pic_save()
    Application.COMAddIns("ESClient10.Connect").Object.saveCase , True, False
End Sub
```

### 拓展(by @heming)
直接获取图片并显示，纯EXCEL实现

```vb
Private Sub Worksheet_Change(ByVal Target As Range)
	Application.EnableEvents = False
	On Error Resume Next
	Shapes.SelectAll
	Selection.ShapeRange.Delete
	Range("A2").Select
	Range("A2").RowHeight = 60  '定义A2的行高，磅数。
	Range("A2").ColumnWidth = 12  '定义A2的列宽，标准字符数。
	Shapes.AddShape(msoShapeRectangle, 0, 24, 72, 60).Select  '定义图片框的左上角位置和宽度、高度。
	Selection.ShapeRange.Fill.Visible = msoFalse
	Selection.ShapeRange.Shadow.Obscured = msoTrue
	Selection.ShapeRange.Shadow.Type = msoShadow18
	Selection.ShapeRange.Fill.UserPicture = "E:\Sys\a.jpg"
	MsgBox ("hello")
	Range("A1").Select
	Application.EnableEvents = True
End Sub
```

### 本节贡献者
*@Castle*  
*@heming* 
