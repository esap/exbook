# 5.8 自动打印指定的打印模板
需求：关于根据不同客户或供应商的要求，自动选择打印相对应模板。
	
### 1.添加字段
在客户或供应商表字段中，添加打印模板字段。这样可以在选择客户时，自己带出打印样式字段  
也可以把打印模板样式表做成独立的表，在填报模板上手工选择对应的打印模板样式  
这里以客户自动带出为例:  
![](./5.8.1.png?raw=true)

### 2.建立打印模板
![](./5.8.2.png?raw=true)
![](./5.8.3.png?raw=true)

### 3.打印代码
```vb
Private Sub CommandButton1_Click()
	Dim str As String
	str = "入库单打印-" & Range("H10").Value
	With Application.COMAddIns("ESClient10.Connect").Object
		.addInitData "MNO", Range("D10").Value
		.newreport str
	End With
End Sub
```

### 4.效果
![](./5.8.4.png?raw=true)
![](./5.8.5.png?raw=true)

### 本节贡献者
*@柳亚子*
