# 4.3 打印检测及打印机相关
## 选择打印机
通过VBA选择指定一个打印机  
```vb
Sub two店()
	On Error Resume Next
	Dim shop As String
	shop = "2店"
	
	Set objRegistry = GetObject("winmgmts:\\.\root\default:StdRegProv")
	objRegistry.GetStringValue &H80000001, "Software\Microsoft\Windows NT\CurrentVersion\PrinterPorts\", shop, a
	pa = shop & " 在 " & Mid(a, InStr(a, ",") + 1, InStr(a, ":,") - InStr(a, ","))
	MsgBox pa
	Application.ActivePrinter = pa
	  
End Sub
```

## 打印检测代码
在 WINXP + EXCEL2003 + ES9.2.335 下正常运行  
```vb
Function 打印机检测()
    Dim net As Object
    Dim Js As Long           '当前电脑已安装打印机计数
    Dim Pt1 As String        '当前默认打印机 含有端口名,如（在 Ne01:)
    Dim Pt2 As String        '当前模打印所需要的打印机名称'如:(Adobe PDF)/(hp LaserJet 1012)
    Dim i As Long            '循环计数
'    Dim Pt3 As String        '当前电脑所有打印机名称
    Dim Pd As Long
    
    打印机检测 = False
       
    Set net = CreateObject("WScript.Network")
'    Set Pts = net.EnumPrinterConnections
    Js = net.EnumPrinterConnections.Count
    If Js < 2 Then
        MsgBox "您的电脑未设置打印机，请到系统--控制面板--打印机和传真--添加打印机！", 48, "系统提示"
        打印机检测 = False
        Exit Function
    End If
    
    Pt1 = Application.ActivePrinter
    Pt2 = "EPSON LQ-730K ESC/P2"
    
   '判定当前电脑是否装有当前模打印所需要的打印机
    For i = 1 To Js - 1 Step 2
        Pt3 = net.EnumPrinterConnections.Item(i)    '打印机名称
        Pd = InStr(Pt3, Pt2)
        If Pd > 0 Then Exit For
    Next
    If Pd = 0 Then
        MsgBox "当前模板需要指定（" & Pt2 & "）打印机打印！您的电脑没有添加此打印机，请先添加此打印后再执行打印作业！", 48, "系统提示"
        打印机检测 = False
        Exit Function
    End If
      
'    '如果所需打印机为当前默认,则直接打印,否则将所需打印机设为默认后再打印
'    If InStr(Pt1, Pt2) Then
''        MsgBox "所需要的打印机已为档前默认！将直接打印"
'        Call 打印 '打印代码
'    Else
''        MsgBox "系统将（" & Pt2 & "）打印机设为默认后再打印！"
'        net.SetDefaultPrinter Pt2         '把默认打印机改为 所需打印机
'        Call 打印  '打印代码
'        Pt1 = Mid(Pt1, 1, Len(Pt1) - 8)     '截去默认打印机端口名（在 Ne01:）
'        net.SetDefaultPrinter Pt1           '还原原先的默认打印机
'    End If
    Set net = Nothing
    打印机检测 = True
    
End Function
```

## 获取打印页数
只能用VBA解决，下面的代码由 `cbtaja` 录制，将下面代码粘贴进 工具--宏--VB编辑器后，以下公式可用：
- =ThisPageNo 显示当前页数，
- =PagesCount 显示总页数；
- =TEXT(ThisPageNo,"第0页 ")&TEXT(PagesCount,"共0页") 在同一单元格显示当前页数和总页数  
```vb
Sub 定义页码及总页数名称()
    ActiveWorkbook.Names.Add Name:="ColFirst", RefersToR1C1:= _
        "=GET.DOCUMENT(61)" '判断打印顺序的设置类型
    ActiveWorkbook.Names.Add Name:="lstRow", RefersToR1C1:= _
        "=GET.DOCUMENT(10)" '本工作表已用到的最大行数
    ActiveWorkbook.Names.Add Name:="lstColumn", RefersToR1C1:= _
        "=GET.DOCUMENT(12)" '本工作表已用到的最大列数
    ActiveWorkbook.Names.Add Name:="hNum", RefersToR1C1:= _
        "=IF(ISERROR(FREQUENCY(GET.DOCUMENT(64),Row())),0,FREQUENCY(GET.DOCUMENT(64),Row()))" 'hNum为本单元格上方的水平分页符个数
    ActiveWorkbook.Names.Add Name:="vNum", RefersToR1C1:= _
                "=IF(ISERROR(FREQUENCY(GET.DOCUMENT(65),Column())),0,FREQUENCY(GET.DOCUMENT(65),Column()))" ''本单元格左边的垂直分页个数
    ActiveWorkbook.Names.Add Name:="hSum", RefersToR1C1:= _
        "=IF(ISERROR(FREQUENCY(GET.DOCUMENT(64),lstRow)),0,FREQUENCY(GET.DOCUMENT(64),lstRow))" ''本工作表最后一个单元格上方的水平分页符个数
    ActiveWorkbook.Names.Add Name:="vSum", RefersToR1C1:= _
                "=IF(ISERROR(FREQUENCY(GET.DOCUMENT(65),lstColumn)),0,FREQUENCY(GET.DOCUMENT(65),lstColumn))" ''本工作表最后一个单元格左边的垂直分页个数
    ActiveWorkbook.Names.Add Name:="ThisPageNo", RefersToR1C1:= _
        "=IF(ColFirst,(hSum+1)*vNum+hNum+1,(vSum+1)*hNum+vNum+1)*ISNUMBER(NOW())" '单元格所在页码
    ActiveWorkbook.Names.Add Name:="PagesCount", RefersToR1C1:= _
        "=GET.DOCUMENT(50)*ISNUMBER(NOW())" '本工作表的总页数
End Sub
```

## 重复打印自动编码并指定打印机
```vb
Sub Macro1()
'ActiveWindow.SelectedSheets.PrintOut Copies:=1, ActivePrinter:="HP LaserJet Professional M1216nfh MFP"
x = InputBox("请输入您要打印的份数")
For i = 1 To x
Cells(3, 2) = i
 Range("A1:F13").Select
    Selection.PrintOut Copies:=1, ActivePrinter:="HP LaserJet Professional M1216nfh MFP", Collate:=True
Next i
End Sub
```

[Excel下载](c4/03/4.3.xls ':ignore')

### 本节贡献者
*@Castle*  
*@kang*   
