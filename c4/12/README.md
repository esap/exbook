# 4.12 不卡不闪的点击变色方案
把数据区域，例如_EST9943定义成_data
 * 支持多行
 * 支持跨行

![](./4.12.jpg)

代码
```vb
Private Sub Worksheet_SelectionChange(ByVal Target As Excel.Range)
    If Intersect(Target, Range("_data")) Is Nothing Then Exit Sub
    On Error Resume Next
    With Range("_data").FormatConditions
        .Delete
    End With
    With Intersect(Target.EntireRow, Range("_data")).FormatConditions
       .Delete
       .Add xlExpression, , True
        .Item(1).Font.ColorIndex = xlAutomatic
        .Item(1).Interior.ColorIndex = 17
    End With
End Sub
```

### 本节贡献者
*@潇湘肥燕*
