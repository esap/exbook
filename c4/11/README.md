# 4.11 VBA应用汇

## VBA指定打印尺寸和其他打印设置
@benava: 重装系统后不想进ES重设模板打印尺寸。

> @ExcelHome

```vb
Sub pripage()    '打印设置
'Application.Dialogs(xlDialogPrinterSetup).Show    '一句话代码，调出系统对话框手动设置选择打印区域
'----------------------------------------------
With ActiveSheet.PageSetup
     '按自定义纸张打印
     '注意：需先在打印设置中自定义一个命名为“SHD”的页面尺寸（长21cm*宽14.7cm）
     .PaperSize = xlPaperSHD       '设置纸张的大小为自定义的“SHD”。若为xlPaperA4则为A4纸     
     .Orientation = xlPortrait        '该属性返回或设置页面的方向。wpsOrientPortrait 纵向；wpsOrientLandscape 横向
     .LeftMargin = Application.InchesToPoints(1.5)
        .RightMargin = Application.InchesToPoints(1.5)
        .TopMargin = Application.InchesToPoints(1.5)
        .BottomMargin = Application.InchesToPoints(1.5)
        .HeaderMargin = Application.InchesToPoints(1)
        .FooterMargin = Application.InchesToPoints(1)
        .PrintGridlines = True
        .CenterHorizontally = True        '页面的水平居中
     '.CenterVertically = True        '页面垂直居中
     .Zoom = False        '将页面缩印在一页内
     .FitToPagesWide = 1        
     If Range("A6") <> "" Then
        .PrintArea = ""    '取消打印区域
        '.PrintArea = "$A$1:$J$21"
        'Range("A1:J21").PrintOut Copies:=1, Collate:=True    '打印指定区域，直接打印
        Range("A1:J21").PrintOut Copies:=1, Preview:=True, Collate:=True   '打印预览。
     End If
        '上面代码即[a1:j21].PrintOut
End With
End Sub
```

## VBA操作网盘上传附件示例
> @lengmu

```vb
oAdd.NFS_uploadFile(pathfrom, pathto, uploadFErr)
Call oAdd.NFS_uploadFile(pathfrom, pathto, uploadFErr)
    If uploadFErr <> "" Then
        
        tishi = tishi & pathfrom & Chr(10) & "附件上传失败！" & Chr(10)
        i = i - 1
        GoTo Loop2
    End If    
Call oAdd.ExecProc("UpMsF", UpMsFErr, Range("T" & m).Value, Range("L4").Value, f, "." + fl)
    If UpMsFErr <> "" Then
        tishi = tishi & pathfrom & Chr(10) & "附件数据更新失败！" & Chr(10)
        i = i - 1
        GoTo Loop2
    End If
Call progressBarShow(i, n)
```

## 点击打开对应文件，例如CAD图纸等(asked by @荆州)
![](./4.11.17.png)

```vb
Sub openF()
    Dim A, Tempfile As String   'Tempfile放CAD文件的路径
    Tempfile = Selection(1).Value
    A = "explorer.exe " & Tempfile
    Call Shell(A, vbMaximizedFocus)
End Sub
```

## NewReport接口如何添加多个数据明细字段
![](./4.15.jpg)

>@范味浓：如果出现错误代码13是字段有合并单元格。

>@村长：真爱生命，远离合并单元格。

## 强行退出所有EXCEL(asked by @执着)
```vb
For Each Process In GetObject("winmgmts:").ExecQuery("select * from Win32_Process where name='EXCEL.EXE'")
    Process.Terminate (0)
Next
```

## 获取桌面路径
Environ("userprofile") & "\Desktop"

## 点击按钮保存表单的宏（注意，使用Application.SendKeys的代码都会锁定小键盘）
Application.SendKeys ("%fs")

## 点击按钮关闭表单的宏
Application.SendKeys ("%fx")

## 点击按钮打开工作台的的宏（07版以后）
Application.SendKeys ("%qr~")

## 点击图片后隐藏图片本身
![](./4.15.1.jpg)

## 点击按钮打开特定超链接
> *@袖子*  
![](./4.11.jpg)

## 纯查询调用存储过程示例(VBA获取当前用户名)
![](./4.15.8.png)

## 动态调用树形或列表
![](./4.15.9.jpg)

## 解决条码、二维码控件连续打印不刷新的问题
![](./4.15.10.jpg)

## 记录模板打印次数
![](./4.15.2.png)

## 打印前设置打印次数
> @BOS-上海-废柴:
比如我要打印，3份，直接在对话框里改一下数字就可以了

> @风云:
```vb
sl = Val(InputBox("请打印数量:", "请打印数量", 1))
PrintOut copies:=sl
```

## 双击弹出对应报表的示例
![](./4.15.3.png)

## 双击切换单元格中的文字
将区域定义为`_data`  
![](./4.15.4.jpg)  
```vb
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
	Cancel = True
	If Intersect(Target, [_data]) Is Nothing Then Exit Sub
	If Target(1) = "Y" Then Target(1) = "N" Else Target(1) = "Y"
End Sub
```

## 检测当前表单打开状态
> *@Kang*  

```vb
If InStr(Application.Caption, "查看") > 0 Then MsgBox "当前是以查看方式打开，不可执行提交作业！", 64, "系统提示": Exit Sub
```

## 检测单元格锁定

```vb
If Range("_ESF42").Locked = True Then MsgBox "当前是以查看方式打开货品档案，不允许执行手动品号作业！", 64, "系统提示": Exit Sub
```

## 关键字快速高亮定位
> *@荊喌*  
![](./4.11.11.jpg)  
```vb
Private Sub CommandButton1_Click()
   AAAA = Empty
    Dim i%, j%, k%, rng As Range
    i = ThisWorkbook.Sheets("Sheet1").Cells(Rows.Count, 2).End(xlUp).Row
    Sheet1.Range("a1:g" & i).Interior.ColorIndex = xlNone
    For Each rng In Range("b2:g" & i)
        If InStr(UCase(rng), UCase(Sheet1.TextBox1)) Then
            rng.Interior.ColorIndex = 6
            If AAAA = Empty Then
            AAAA = rng.Address
            End If
        End If
    Next
   Range(AAAA).Activate
End Sub
```

[Excel下载](c4/11/4.11.11.xls ':ignore')

## 数据透视表刷新
```vb
ActiveSheet.PivotTables("数据透视表1").PivotCache.Refresh
```

## VBA控制字段必填示例
> *@Kang*  
```vb
Function Data检测()
    Dim MiniRow As Integer
    Dim MaxiRow As Integer
    Dim MaxiCol As Byte
    Dim MyRow As Byte
    Dim MyCol As Byte
    Dim myRange As String    
    Djs = Application.WorksheetFunction.CountA(ActiveSheet.Range("_ESF255"))
    If Djs < 1 Then
        MsgBox "至少填写一个问题点，请填写数据！", 64, "系统提示"
        Range("C13").Select
        Data检测 = False
        Exit Function
    End If
    Data检测 = False
    MiniRow = 13
    MaxiRow = MiniRow + Range("_ESF255").Rows.Count - 1    
    If Range("C" & MaxiRow) = "" Then
        MaxiRow = Range("C" & MaxiRow).End(3).Row()
    Else
        Djs = Application.WorksheetFunction.CountA(ActiveSheet.Range("_ESF2971"))
        If Djs >= 2 Then
            MaxiRow1 = Range("_ESF2971").End(xlDown).Row()
        End If
    End If    
    If MaxiRow < 13 Then
        MsgBox "当前没有数据，请填写完整！", 16, "系统提示"
        work数据检测 = False
        Exit Function
    End If    
    MaxiCol = 5    
    For MyRow = MiniRow To MaxiRow
        For MyCol = 3 To MaxiCol Step 1
            If Cells(MyRow, MyCol).Value = "" Then
                myRange = Cells(MyRow, MyCol).Address(RowAbsolute:=False, ColumnAbsolute:=False)
                MsgBox " 检测到数据区域内【" & myRange & "】栏数据未填写完整，请在选中栏内填写数据后再提交！", 32, "系统提示"
                Cells(MyRow, MyCol).Select
                Data检测 = False
                Exit Function
            End If
        Next MyCol
    Next MyRow    
    Data检测 = True
End Function
```

## 限制重复打开查询模板代码
> *@Kang*  
```vb
'工作簿代码
Private Sub Workbook_Open()
    Call EnumWindows(AddressOf 重复窗口, 0&)
End Sub
'模块代码
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Function 重复窗口(ByVal lhWnd As Long, ByVal lParam As Long) As Long
    Dim Title21 As String * 255                                                 '窗口标题
    窗口名称 = "填报：SCRB_1"    
    Call GetWindowText(lhWnd, Title21, 255&)                                    '获取窗口标题和窗体句柄
    If (InStr(Title21, Chr(0&)) > 0&) Then
        窗口名称2 = Left(Title21, InStr(Title21, Chr(0&)) - 1&)
        If InStr(窗口名称2, 窗口名称) > 0 Then
            'MsgBox "有重复"
'            For i = 1 To Windows().Count
'                If Windows(i).Caption = 窗口名称 Then Windows(i).Activate
'            Next            
            ThisWorkbook.Close
            ThisWorkbook.Close
            Exit Function
        End If
    End If
   重复窗口 = True
End Function
```

## 延时程序
```vb
'调用时，使用 delay 3  即可延时3秒
sub delay(T as single)
	dim T1 as single
	t1=timer
	do
		doevents
	loop while timer-t1<t
end sub
```

应用示例：每隔3秒执行指定表间公式

![](./4.11.aotudo.jpg)

## 判断单元格是否属于某区域
```vb
'方法一
If Union(ActiveCell, Range("aa")).Address = Range("aa").Address Then
	MsgBox "包函"
Else
	MsgBox "不包函"
End If
'方法二
If Not Intersect(ActiveCell, Range("aa")) Is Nothing Then
	MsgBox "包函"
Else
	MsgBox "不包函"
End If
```

## 执行表间公式前先弹出确认对话框
> 挨踢熊
```vb
Sub 按钮1_Click()
Dim t
t = MsgBox("取数有风险，是否继续？", vbYesNo + vbQuestion, "挨踢熊")
If t = vbYes Then
    Dim oAdd As Object
    Set oAdd = Application.COMAddIns("ESClient10.Connect").Object
    oAdd.execquery ("查货品,查期初,查出入")
    Set oAdd = Nothing
End If
End Sub
```

## 一句话VBA(ES-VBA接口)
```vb
'-----------------------------------------------------
'VBA 简化代码 by woylin 2014-9-28
'-----------------------------------------------------
'★提数★
Private Sub btn1_Click()
	 Application.COMAddIns("esclient10.connect").Object.execquery "提数 1"
End Sub
'★提数（带弹出提示）★
Private Sub btn1_Click()
	if MsgBox("确认执行提数？", 1+64)=1 then Application.COMAddIns("esclient10.connect").Object.execquery "提数 1"
End Sub
'★回写★
Private Sub btn2_Click()
	 If Application.COMAddIns("esclient10.connect").Object.execupdate("回写 1") Then MsgBox "OK"
End Sub
'★新建★
Private Sub btn4_Click()
	 Application.COMAddIns("esclient10.connect").Object.newreport "入库单"
End Sub
'注意第二参数用于控制关闭调用源表单，官方教程中未指出
Private Sub btn5_Click()
	 Application.COMAddIns("esclient10.connect").Object.newreport "出库单", 1
End Sub
'★弹出规范★
Private Sub btn1p_Click()
	 [i6].Select: Application.COMAddIns("esclient10.connect").Object.poptree "名字_树"
End Sub
'注意该接口可调用列表规范，官方教程中未指出
Private Sub btn2p_Click()
	 [d6].Select: Application.COMAddIns("esclient10.connect").Object.poptree "批次库存_列表"
End Sub
'★插入行★
Private Sub btn9_Click()
	 Application.COMAddIns("esclient10.connect").Object.InsertRow 1, 6, 1
End Sub
'★保存★
Private Sub btnSave_Click()
	 Application.COMAddIns("esclient10.connect").Object.savecase , , 0
End Sub
'★存储过程★
Private Sub btn1n_Click()
	 Dim sErr$
	 If Application.COMAddIns("esclient10.connect").Object.ExecProc("p_2", sErr, "") Then MsgBox "已更新"
End Sub
'★存储过程(带结果集)★
Private Sub btn2n_Click()
	 Dim Rs As New ADODB.Recordset, sErr$
	 If Application.COMAddIns("esclient10.connect").Object.ExecQryProc("p_1", Rs, sErr, "") Then myGrid1.SetDatasource Rs
End Sub
```

## 通过扫描仪连续执行扫描的例子
![](./4.11.gif)

[Excel下载](c4/11/4.11.1.xls ':ignore')


## 从sql读取数据并显示的例子
直接用VBA从数据库取数，纯EXCEL也能做到，不过数据库连接参数会暴露在代码中，不如用ES存储过程接口安全

[Excel下载](c4/11/4.11.2.xls ':ignore') 

## 从sql读取图片数据流并显示的例子
如果ES不开启网盘，图片附件会以二进制方式存储在数据库中，与普通读取方式不同的是图片附件的读取为“流式”读取

<a href="../src/4.11.1.xls" download>Excel下载</a>  

## 纯查询必填检查示例
![](./4.15.5.jpg)  
```vb
'按钮宏
Sub xSave()
If xCheck(Range("d7,d10:d12")) Then Exit Sub  'd7,d10:d12改成需要验证的单元格区域
'TODO: 这里是验证通过后的正常代码
MsgBox "验证通过！"
End Sub
'验证函数 by woylin 2015.3.21
Function xCheck(rng As Range) As Boolean
    Dim cel
    For Each cel In rng
        If cel.Value = "" Then
            MsgBox cel.Offset(0, -1) & "必填(" & cel.Address & ")"
            xCheck = True
            Exit Function
        End If
    Next
    xCheck = False
End Function
```

## 生成时间明细
![](./4.15.6.jpg)    
```vb
Private Sub CommandButton1_Click()
	a = Range("A65536").End(xlUp).Row
	
	If CommandButton1.Caption = "生成" Then
		CommandButton1.BackColor = &HFF8080
		For h = 2 To 100
			'起始时间是9:12 增加量是2分27秒
			Cells(h, 1) = Evaluate("Time(9, 12 + 4*" & h - 2 & ", 27*" & h - 2 & ")")   
			Cells(h, 1).NumberFormatLocal = "h:mm;@"  '格式
		Next
		CommandButton1.Caption = "清除"
		CommandButton1.BackColor = &HFF&
	Else
		Range("A2:A" & a).Clear
		CommandButton1.Caption = "生成"
		CommandButton1.BackColor = &HFF8080
	End If
End Sub
```

## 生成日期明细
> @荆州
![](./4.11.16.png) 

## 表内自定义公式
![](./4.15.7.jpg)    
```vb
'thisworkbook中的代码
Private Sub Worksheet_Change(ByVal Target As Range)
    Cells.Calculate
End Sub
'插入一个模块，填入下面函数代码
Function xEval(rng As Range, rngs As Range)
    '第一个参数是公式字符串，第二个是参数表
    Dim tmp$
    tmp = rng.Value
    '遍历参数表进行替换
    For i = 1 To rngs.Count
        tmp = Replace(tmp, "{" & i & "}", CSng(rngs.Item(i).Value))
    Next
    'TODO： 这个函数类似C语言的printf,使用{1}，{2}。。。这样的文本格式定义公式
    xEval = Application.Evaluate("=" & tmp)
End Function
```

## 直接叫出工作台   
```vb
'2007版无效，适用于03和10版
Sub xxx()
     Application.CommandBars("Worksheet Menu Bar").Controls("报表(&R)").Controls(1).Execute
End Sub
```

## 使用ADO直接连接数据库又不暴露连接参数的一种思路
> @Kang：提取加密字串到本地，通过自定义函数Decrypt()解析  
```vb
   '建立与指定SQL Server数据库的连接
    If Sheet5.Range("B3").Value = "" Or Sheet5.Range("B4").Value = "" Then Exit Sub
    If Sheet5.Range("B2").Value = "启用" Then strcnn = Sheet5.Range("B3").Value
    If Sheet5.Range("B2").Value = "测试" Then strcnn = Sheet5.Range("B4").Value
    If strcnn <> "" Then
        strcnn = Decrypt(strcnn)
        cnn.ConnectionString = strcnn
    Else
        Exit Sub
    End If
    cnn.Open
```

## 当前sheet另存到本地示例   
```vb
Sub xSave()
    Dim n%, a%
    Dim wb, wb2 As Workbook
    Dim i%
    Application.ScreenUpdating = False
    Set wb = ThisWorkbook.ActiveSheet
    Set wb2 = Workbooks.Add
    ActiveWindow.DisplayGridlines = False
'    For i = 1 To wb.Sheets.Count
        wb.Cells.Copy wb2.Sheets(1).Cells
'    Next
'    wb2.Sheets(1).Shapes("XSAVES").Delete
	'**********以下代码调整格式用，请忽略*************'
    With wb2.Sheets(1)
        n = 8
        i = 1
        Do
            .Cells(n, 1) = i
'            .Cells(n, 1).RowHeight = 14.25
            .Cells(n + 1, 1).EntireRow.Insert
            .Cells(n + 1, 2) = "    订单特殊要求：" & .Cells(n, 9).Value
            .Cells(n, 9) = ""
            With .Range("B" & n + 1 & ":G" & n + 1)
                .Merge
                .Validation.Delete
                .HorizontalAlignment = xlGeneral
                a = Application.Max(Len(.Item(1).Value) / 40, Len(.Item(1).Value) - Len(Application.Substitute(.Item(1).Value, VBA.Chr(10), "")) + 1)  '25根据实际情况自己调整
                Rows(.Row).RowHeight = 14.25 * (a + 1)
            End With
            n = n + 2
            i = i + 1
        Loop While (.Cells(n, 1).Value <> "交货条件")
        .PageSetup.PrintArea = "$A:$G"
    End With
	'**********以上代码调整格式用，请忽略*************'
    Application.ScreenUpdating = True
    Set wb = Nothing
    Set wb2 = Nothing
End Sub
```
