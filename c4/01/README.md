# 4.1 VBA明细选择判定示例
* 效果图：  
![](./4.1.1.jpg?raw=true)

* 实现在Excel表格中点击不同明细区域时，区域首行更新为选中行的数据  
```vb
Private Sub Worksheet_SelectionChange(ByVal Target As Range)    
    Dim rng
    Set rng = Application.Intersect(Target(1).EntireRow,Range("_data1")) '检查是否为data1行区域    
    If rng Is Nothing Then
        Set rng = Application.Intersect(Target(1).EntireRow,Range("_data2")) '检查是否为data2行区域
        If rng Is Nothing Then Exit Sub
        Range("_check2").Value = rng.Value '行拷贝
    Else
         Range("_check1").Value = rng.Value '行拷贝
    End If
    Set rng = Nothing	
End Sub
```

### 本节贡献者
*@DHL*
 