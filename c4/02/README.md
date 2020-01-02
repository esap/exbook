# 4.2 提取中英文字符
* 效果图：  
![](./4.2.jpg?raw=true)

```vb
Function getchn(MyValue As Range)	'getchn获取中文
	Dim i As Integer
	Dim chn As String
	For i = 1 To Len(MyValue)
		If Asc(Mid(MyValue, i, 1)) < 0 Then
		chn = chn & Mid(MyValue, i, 1)
		End If
	Next
	getchn = chn
End Function
	
Function yw(str As String)			'yw获取英文
	 With CreateObject("vbscript.regexp")
	     .Global = True
	     .Pattern = "[^a-zA-Z]"
	     yw = .Replace(str, " ")
	End With
End Function
```

- 另一个版本<a href="../src/4.2.xls" download>Excel下载</a>

### 本节贡献者
*@heming*
