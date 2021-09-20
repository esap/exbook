# 5.7 明细行锁定-审核记录不能修改删除
### 另一个新版本
```vb
Private Sub Worksheet_SelectionChange(ByVal Target As Range) '单元格选择事件
	Dim r As Integer, p As String, v As Single
	r = ActiveCell.Row '赋值变r = 当前选定单元格的行号
	p = Cells(r, 29).Value  '赋值变p = 当前选定单元格所在行的第29列（状态列）的单元格的值
	v = Application.Version  '赋值变量v = 当前Excel应用程序的版本号	
	If p = "关闭" Or p = "审核" Then    '如果状态为关闭或审核
	    Application.ExecuteExcel4Macro "show.toolbar(""ribbon"",false)" '屏蔽功能区
	    Application.CommandBars("Ply").Enabled = False    '屏蔽右键
	    Application.CommandBars("cell").Enabled = False   '屏蔽右键
	  If v = "11.0" Then  '如果版本为Excel2003
	   Application.CommandBars("表单操作").Visible = False  '屏蔽表单操作
	  End If
	  MsgBox "不允许修改和删除状态为关闭或审核的数据，因为这会导致金蝶系统异常"
	End If	
	If p <> "关闭" And p <> "审核" Then    '如果状态不为关闭并且不为审核
	    Application.ExecuteExcel4Macro "show.toolbar(""ribbon"",true)" '恢复功能区
	    Application.CommandBars("Ply").Enabled = True  '恢复右键
	    Application.CommandBars("cell").Enabled = True  '恢复右键
	  If v = "11.0" Then '如果版本为Excel2003
	   Application.CommandBars("表单操作").Visible = True '恢复表单操作
	  End If
	End If
End Sub
```

- 资料下载：

@清风版：[模板下载](c5/07/5.7.5.zip ':ignore')

@荆州版：[模板下载](c5/07/5.7.6.zip ':ignore')

### 老版本
- 效果图:

![](./5.7.1.png?raw=true)

@荆州版：[模板下载](c5/07/5.7.2.rar ':ignore')

- 老老方案Excel版：

版本一：@荆州版：[模板下载](c5/07/5.7.3.xls ':ignore')

版本二：[模板下载](c5/07/5.7.4.xls ':ignore')

### 补充
* 2016版测试OK

### 本节贡献者
*@昆明haotian*  
*@清风*  
*@荆州*  
