# 5.1 使用VBA+SQL存储过程代替提数

## 第一步：引用ADO2.8
![](./5.1.1.png?raw=true)
 
插入模块，然后把代码复制进模块执行

## 第二步：VBA代码
```vb
Sub 查询()   '这是主执行sub
	Dim Rs As New ADODB.Recordset, sErr$, r As Integer
	Call delrow
	With Application.COMAddIns("esclient10.connect").Object
		.ExecQryProc "V_成品入库明细", Rs, sErr, Range("开始日期").Value, Range("结束日期").Value
		r = Rs.RecordCount
		.InsertRow 1, 10, r - 2
	End With
	Range("明细").CopyFromRecordset Rs
End Sub

Sub delrow()  '这是子执行sub，用于删除行和清除内容操作
	Dim saddr$
	Dim saddr2$
	Dim saddr3$
	Dim a As Long
	Dim b As Long
	Dim c As Long
	Dim d As Long
	Dim i As Long
	Dim rng As Range
	With Application.COMAddIns("esclient10.connect").Object
		.GetFieldAddress "公司订单号", saddr, a, b, c, d
		.GetFieldAddress "物料流水码", saddr2
		i = c - a
		saddr3 = VBA.Left(saddr, InStr(saddr, ":")) & VBA.Right(saddr2, Len(saddr2) - InStr(saddr2, ":"))
		Range(saddr3).ClearContents
		.deleteRow 1, 10, i - 1
	End With
End Sub
```

## 第三步：SQL存储过程
```sql
create proc [dbo].[V_成品入库明细]
@开始日期 date,
@结束日期 date
as
select
	a.公司订单号,b.物料编码,b.规格,b.物料名称,a.发生数量,a.记帐方向,a.摘要,a.物料流水码
from
	(
		select
			a.公司订单号,
			a.物料流水码,
			发生数量=sum(isnull(a.数量,0)),
			记帐方向=1,
			摘要='成品入库'
		from
			成品入库单_明细 a
		inner join
			成品入库单_主表 b
		on
			a.ExcelServerRCID=b.ExcelServerRCID
		where
			ISNULL(b.作废,'')<>'是' and b.入库日期>=@开始日期 and b.入库日期<=@结束日期
		group by
			a.物料流水码,a.公司订单号
	) a
left join 物料表 b
on a.物料流水码=b.物料流水码
```

## 第四步：单元格区域需要用名称定义，图例
![](./5.1.2.png?raw=true)

### 资料包下载
<a href="./5.1.zip" download>模板下载</a>

## 延伸阅读-如何使用ES存储过程接口
```vb
'******************************
'在运行这个例子之前，首先在 ESSys 数据库中建立如下的两个存储过程：
'create proc p_test1(@a int)
'as
'begin
'    select @a*2
'End

'create proc p_test2(@a int)
'as
'begin
'    -----可能的操作
'    Return
'End

'因为这里第一个存储过程要返回记录集，所以，还需要引用 "Micorsoft ActiveX Data Objects 2.8"，如果存储过程不需要返回记录集，则不需引用它
'******************************

Private Sub CommandButton1_Click()
    Dim oAdd As New ESClient.Connect
    Dim a As Integer, b As Integer
    Dim sErr As String
    Dim rs As New ADODB.Recordset
    
    a = Range("B2")
    
    '这个存储过程是返回记录集的
    oAdd.execQryProc "p_test1", rs, sErr, a
    If Not rs.EOF Then
        MsgBox rs(0)
    End If
    Set rs = Nothing
    
    '这个存储过程是不返回记录集的
    oAdd.execProc "p_test2", sErr, a
    
End Sub
```

### 本节贡献者
*@毛毛*  
*@清风*
