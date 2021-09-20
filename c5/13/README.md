# 5.13 ES调用存储过程的后处理-提数填充工具
### **目的**
ES调用存储过程并提数的方法和实现已经有人做教程说明过了，不再赘述。但是在实际应用中，多次执行查询存储过程，又涉及到上一次结果的清除，行数的重新分配，如果行列不匹配还要循环recordset对应等等细节，每次都要不厌其烦的去做重复的工作，所以，分享一个我自己在用的工具模块，旨在便捷调用存储过程后一键填充，接近调用原生提数公式一样的体验～

### **先上调用方法和使用场景：**
-  ###调用前准备：###

   + 下载附件 CopyFromProcResult.bas 
   + 在需要调用存储过程提数的模板，进入vba编辑器，左边项目中右键--导入文件，选择CopyFromProcResult.bas后导入。
   + 记得在工具-引用 ：Microsoft ACtiveX Data Object XX版本。因为用到的Recordset需要这个引用
   + 模板页面定义好需要填充的重复数据区域。

-  ###场景一：###
   存储过程执行后，得到结果集，直接填充到excel，明细单元格定义的行列和存储过程得到的结果集一一对应。
``` vb
Dim oAdd As Object
Dim rs As ADODB.Recordset
Dim errMsg As String
Set oAdd = Application.COMAddIns("ESClient10.Connect").Object
If oAdd.ExecQryProc("p_QueryBOMReverse", rs, errMsg, Range("C3")) = False Then
   MsgBox "Exec Proc Error:p_QueryBOMReverse proc" + errMsg
   Exit Sub
End If
Set oAdd = Nothing
```
以上部分是ES官方指导的存储过程调用方式
以下是调用工具模块
``` vb
DirectCopyRs Sheet1, "_EST7", rs
```

__参数说明__
	"Sheet1"--填充区域所在sheet页名称 
	"_EST7"--要填充的区域名称（what？这个在哪里？用鼠标选定之前定义好的重复数据区域，看屏幕左上角.....）
总之一句话搞定，包括多次运行的数据行处理也自动完成
-  ###场景二：###
   存储过程执行后，得到结果集，和excel中定义的列并不是完全对应。
``` vb
Dim oAdd As Object
Dim rs As ADODB.Recordset
Dim errMsg As String
Set oAdd = Application.COMAddIns("ESClient10.Connect").Object
If oAdd.ExecQryProc("p_QueryBOMReverse", rs, errMsg, Range("C3")) = False Then
   MsgBox "Exec Proc Error:p_QueryBOMReverse proc" + errMsg
   Exit Sub
End If
Set oAdd = Nothing
```
以上部分是ES官方指导的存储过程调用方式，好吧，这个是复制上面一段的。 
以下调用就厉害了：
```vb
Dim col As Collection
Set col = New Collection
col.Add "B:Customer"
col.Add "E:CustomerPdtName"
col.Add "F:TreeMID"
col.Add "G:MPdtName"
col.Add "I:Qty"
col.Add "J:mPdtAttr"
MapCopyRs Sheet1, "_EST7", rs, col
```

使用方法就是定义个Collection，然后添加Excel列号和存储过程返回结果集字段的映射.

B:Customer 意思是 B列绑定 Customer这个字段。 中间是冒号":"隔开。 

最后调用下MapCopyRs，省却了循环Recordset的诸多烦恼～yeah。

**有了这个工具，在超复杂的SQL语句摆在面前时，就不再纠结于是用各种辅助提数到晕倒，还是去使用存储过程后痛苦的处理页面填充了^_^  **

最后是工具模块的超链接：[CopyFromProcResult.bas](c5/13/CopyFromProcResult.bas ':ignore')

### 本节贡献者
@Meteor
2017-1-3
