# ESToolKit API

### 准备工作
1. vba工程中导入<a href="./ESToolKit.bas" download>EsToolKit.bas</a>
2. 将<a href="./p_ESTK_FilterInfo.sql" download>p_ESTK_FilterInfo.sql</a>执行入数据库
3. vba工程引用 Microsoft Scripting Runtime

### GetF(FieldName as String)
```vb
'描述: 根据字段名（数据项定义的名称）得到对应的值
'注意: FieldName格式： 数据表表名.字段名
'      仅用于单一数据项，如果是要得到重复数据项的值，请用GetF_dt
'调用示例：
	Dim curStepNo As String
	curStepNo = GetF("WP_Process.CurrentSetNo")
	'Memo：WP_Process为定义的表名，CurrentSetNo为字段名
```
### GetF_Dt(FieldName As String, rowIndex As Long)
```vb
'描述: 由字段名（FieldName）和 数据行号（rowIndex）得到对应的值
'注意: rowIndex 为数据行号，即得到本次明细项目的第N行的数据，而不是在excel中的绝对行坐标
'     如果是excel的当前行号，可用GetRelateLine函数来转化成数据行号
'调用示例：
	Dim StepName As String
	StepName=GetF_Dt("WP_Step.StepName",2)
	'Memo: 得到WP_Step表中第二行的数据的StepName的值
	StepName=GetF_Dt("WP_Step.StepName",GetRelateLine(13,"WP_Step"))
	'Memo：如果已知目标行在excel中的绝对行号是13，则用GetRelateLine得到数据行号（RelateLine）
```
### SetF(FieldName as String,val)
```vb
'描述: 向对应的字段名赋值（数据项定义的名称）
'注意: FieldName格式： 数据表表名.字段名
'      仅用于单一数据项，如果是要给重复数据项赋值，请用SetF_dt
'调用示例：
	SetF("WP_Process.StepName","电焊")
```
### SetF_Dt(fieldName As String, rowIndex As Long, val)
```vb
'描述: 向数据行（rowIndex）的字段名（FieldName）赋值
'注意: rowIndex 为数据行号，即得到本次明细项目的第N行的数据，而不是在excel中的绝对行坐标
'     如果是excel的当前行号，可用GetRelateLine函数来转化成数据行号
'调用示例：
	Dim StepName As String
	SetF_Dt("WP_Step.StepName",2,"电焊")
```
### GetRelateLine(absLine As Long, tableName As String) 
```vb
'描述: 将excel绝对行号(absLine)转化成数据行（RelateLine)
'参数: tableName--模板中定义的表名
'返回值：Long
```
### GetAbsLine(relLine As Long, tableName As String) 
```vb
'描述: 将数据行（RelateLine)转化成excel绝对行号(absLine)
'参数: tableName--模板中定义的表名
'返回值：Long
```
### GetTableRange(tableName As String, ByRef shtIndex As Integer)
```vb
'描述: 由表名tableName得到该表当前定义的Range
'参数: tableName--模板中定义的表名
'     shtIndex---用来回传所在的sheet的索引号
'返回值：Range
```
### GetFRange(fieldName As String)
```vb
'描述: 由列名得到该列当前定义的Range
'参数: fieldName--模板中定义的表名
'返回值：Range
```
### GetFAddr(fieldName As String)
```vb
'描述: 由字段名得到地址值
'返回值：String
	Dim vAddr as String
	vAddr=GetFAddr("WP_Process.StepName") 
	'得到"$A$2:$B$2"
```
### Focus(fieldName As String)
```vb
'描述: 定位到定义字段的位置，如果定义的字段在其他sheet，会自动跳转到对应的sheet
	Focus "WP_Process.StepName"
```
### AddRows(tableName As String, rowCount As Long)
```vb
'描述: 往表中(重复数据项)末尾添加N行，N即参数rowCount，表名：tableName
	AddRows "WP_Process_Dt",3
```
### ClearRows(tableName As String)
```vb
'描述: 清空数据表的所有行，表名：tableName
	ClearRows "WP_Process_Dt"
```
### DelOneRow(tableName As String, recordIndex As Long)
```vb
'描述: 删除表的某一行，表名：tableName，recordIndex为数据行号（请参考上述数据行和绝对行的区别描述）
	DelOneRow "WP_Process_Dt",4
```
### DelRowsByFilter(tableName As String, filterDict As Dictionary)
```vb
'描述: 根据筛选条件（filterDict），对数据表进行删除符合条件的行
'filterDict: key--数据字段名FieldName（带表名）,value--"对应的值"
'注意 Microsoft Scripting Runtime的引用
	Dim fdict as Dictionary
	set fdict=new Dictionary
	fdict.Add "WP_Process_Dt.StepName","电焊"
	fdict.Add "WP_Process_Dt.WorkerType","A"
	DelRowsByFilter "WP_Process_Dt",fdict
	'删除全部 StepName="电焊" And WorkerType="A" 的记录
```
### getRowCount(tableName As String) as Long
```vb
'描述: 得到数据表全部行数（包含空行）
```
### getLastRow(tableName As String) As Long
```vb
'描述: 得到数据表的最后一行（包含空行）
```
### getFirstRow(tableName As String) As Long
```vb
'描述: 得到数据表的第一行（包含空行）
```
### IsDesign() as Boolean
```vb
'判断模板是否在设计期，True--在设计期，False--在使用期
'一般用于Worksheet_SelectionChange等事件处理函数中，在设计时不想被频繁的点击而触发，在使用时再触发
```