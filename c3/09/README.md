# 3.9 手工SQL添加es数据记录
需求：需要用存储过程(或者其他程序代码)往es库添加记录，并能在工作台中显示出来.

## 分析  
如果使用工作台填报某个模板，保存的记录会带上es的一些标记，比如ExcelServerRCID，RN。如果存在工作流的话还会增加ExcelServerWIID等等，另外还会对系统的一些维护表进行写入数据。所以，如果直接用SQL的一条INSERT语句插入的数据只会出现在后台，不会出现在工作台的显示界面。
	
本节就最简单的情况做个解决方案。

## 适应条件
* 模板(数据表)中没有用自动编号数据规范的字段;  
* 没有使用工作流

## 案例步骤

* 准备表格:B_Department 部门表(自定义的,不是系统自带的).就一个有效字段Department.  

```sql
select * from B_Department
```
![](./3.9.1.png)

* 可以用存储过程来实现,以下就显示sql的代码:  

```sql
--增加自增编号(和系统的同步)
exec GetNewId_s 26,1   
	declare @rcid varchar(20)
declare @maxid int
select @maxid=maxId from ES_SysId_s where IdName=26 and idDate=convert(varchar(10),getdate(),111)
set @rcid='rc'+convert(varchar(10),getdate(),112)+ REPLICATE('0',5-len(@maxid))+convert(varchar(5),@maxid)
  --用INSERT语句插入主数据
insert into B_Department(Department,ExcelServerRCID,ExcelServerRN, ExcelServerCN,ExcelServerRC1, ExcelServerWIID,ExcelServerRTID,ExcelServerCHG)
values('xxx',@rcid,1,0,'','',36.1,0)
--ExcelRCID:生成的自增ID
--ExcelServerRN:一般是明细表用,1,2等增加,主表为1即可
--ExcelServerCN,ExcelServerRC1, ExcelServerWIID 按照以上默认值设置即可,
-- ExcelServerRTID是模板编号,插入哪个模板,就固定哪个编号.
-- ExcelServerCHG,按默认0输入. 
--在[ES_RepCase]中注册这条记录
INSERT INTO [PAPS].[dbo].[ES_RepCase]
   ([rcId],[RtId],[fillDept],[fillDeptName],[fillUser],[fillUserName],[fillDate]
,[rcDesc],[state],[lstFiller],[lstFillerName],[lstFillDate],[backUpdate],[openState],[openBy],[openByName],[OpenBySesId],[lockState],[lockInServer],[noticeState],[setNStateInServer],[replacerId_fill],[replacerName_fill],[replacerId_lstFill],[replacerName_lstFill],[printTime],[wiId],[commitByDataWriter])
VALUES
           (@rcid    --rcid   生成的RCID
           ,36.1 --<RtId, nvarchar(20),>  模板号
           ,2--<fillDept, int,>  填报部门号(系统表)( 在ES_Dept表中)
           ,''--<fillDeptName, nvarchar(50),>  填报部门名(在Es_Dept表中)
           ,1--<fillUser, int,> 填报用户号(在ES_User中的UserID)
           ,'Admin'--<fillUserName, nvarchar(50),> 填报用户名(在ES_User中的UserID)
           ,getdate()--<fillDate, datetime,>--当前日期
           ,''--<rcDesc, nvarchar(2000),>
           ,1--<state, smallint,>
           ,1--<lstFiller, int,>
           ,'Admin'--<lstFillerName, nvarchar(50),>
           ,getdate()--<lstFillDate, datetime,>--当前日期
           ,0--<backUpdate, smallint,>
           ,0--<openState, smallint,>
           ,NULL--<openBy, int,>
           ,NULL--<openByName, nvarchar(50),>
           ,NULL--<OpenBySesId, nvarchar(20),>
           ,0--<lockState, smallint,>
           ,0--<lockInServer, smallint,>
           ,0--<noticeState, smallint,>
           ,0--<setNStateInServer, smallint,>
           ,0--<replacerId_fill, int,>
           ,''--<replacerName_fill, nvarchar(50),>
           ,0--<replacerId_lstFill, int,>
           ,''--<replacerName_lstFill, nvarchar(50),>
           ,NULL--<printTime, datetime,>
           ,''--<wiId, nvarchar(20),>
           ,0--<commitByDataWriter, smallint,>
)
--以上为添加的结构,需要插入的如当前人当前部门等按照实际业务需要插入,就比较灵活了.
```
涉及到工作流的待续，另外在获取RCID等可能相关的数据表:   
	原表名_Wi ,
	ES_WorkItem    ES_WfCase   ES_Witodo
	ES_IdUsed--这个用于有自动编码规范的

## 简易用法示例
* 在es中新建个模板，Student,  学生信息表  

就三个字段   **姓名**  **国家**  **城市**

在工作台用页面新增的方式先插入一条记录
 
* 进入sqlserver查询分析器

* 先将存储过程的脚本DoAfterInsert.sql 在目标库执行下，将这个存储过程创建出来

* 用sql模拟执行一条插入语句：
```sql
INSERT INTO dbo.Student 
        ( 姓名 ,
          国家 ,
          城市 
        )
VALUES  ( N'王五' , -- 姓名- nvarchar(20)
          N'美国' , -- 国家- nvarchar(20)
          N'纽约' -- 城市- nvarchar(20)
        )
```
* 这时候Student这张表有了记录，但是如果回到es工作台刷新下，这条记录是没有显示出来的。

* 在sqlserver查询分析器执行存储过程：

EXEC DoAfterInsert 'Student','27.1',1,1,' and 姓名=''王五'''

参数说明：

?> `Student`是要插入的表名，  
`27.1`是模板号，（可以select * from Student ，找到从es工作台插入的张三这条记录的ExcelServerRTID字段一样的，先插入条记录也是为了这个目的。）  
接下来的两个`1`，其实一个是用户号，一个是部门号，默认是1号用户，也就是admin，1号部门，也就是默认部门插入的。如果有自定义需求，可以修改，去匹配系统的部门和用户表得到各自id即可。  
` and 姓名='王五'` 这个是定位语句，即用where条件语句定位到刚插入的这条记录的唯一条件。注意sql引号转义的用法


SO，总结下，如果是外挂程序插入记录，要遵循

插入一条记录，执行DoAfterInsert一下，

插入一条记录，执行DoAfterInsert一下，

插入一条记录，执行DoAfterInsert一下，

。。。。。。。 

这样的循环逻辑。


## 总结
利用以上原理,用一个存储过程DoAfterInsert(参附录),在插入数据后执行这个存储过程即可.

存储过程参数说明:   
* @tableName---表名,
* @rptId--模板序号 
* @userID--用户号   
* @DeptID---部门号  
* @where--插入数据的定位条件,注意要以" and"开头

调用示例(注意引号的转义):
```sql
insert into B_Department(Department) values('okok')
exec DoAfterInsert 'B_Department','36.1',1,1,' and Department=''okok'''
```

## 附录一（DoAfterInsert存储过程）
```sql
	create proc DoAfterInsert(@tableName varchar(50),@rptId varchar(20),@userID int ,@deptID int ,@where varchar(1000))
	as
	begin
	--declare @tableName varchar(50)  --param
	--declare @rptId varchar(20)   --param
	--declare @userID int  --param
	--declare @deptID int --param
	--declare @where varchar(1000)--param
	--set @tableName='B_Department'
	--set @rptId='36.1'
	--set @userID=1
	--set @DeptId=1
	--set @where =' and B_Department=''xxx'''

	declare @sqlexec nvarchar(4000)
	declare @userName varchar(50)
	declare @DeptName varchar(50)


	--1.ExcelServer的最重要的RCID编号递增
	exec GetNewId_s 26,1
	declare @rcid varchar(20)
	declare @maxid int
	select @maxid=maxId from ES_SysId_s where IdName=26 and idDate=convert(varchar(10),getdate(),111)
	set @rcid='rc'+convert(varchar(10),getdate(),112)+REPLICATE('0',5-len(@maxid))+convert(varchar(5),@maxid)

	set @sqlexec='update '+@tableName+' set ExcelServerRCID='''+@rcid+''',ExcelServerRN=1,ExcelServerCN=0,ExcelServerRC1='''',ExcelServerWIID='''',ExcelServerRTID='''+@rptId+''',ExcelServerCHG=0
	  where 1=1 '+@where
	--print @rcid
	 exec(@sqlexec)

	select @userName=UserName from ES_User where userID=@userID
	select @DeptName from ES_Dept where DeptId=@DeptId

	INSERT INTO [dbo].[ES_RepCase]
	([rcId] ,[RtId],[fillDept],[fillDeptName],[fillUser],[fillUserName],[fillDate],[rcDesc] ,[state] ,[lstFiller] ,[lstFillerName],[lstFillDate],
	[backUpdate],[openState],[openBy] ,[openByName] ,[OpenBySesId],[lockState] ,[lockInServer],[noticeState],[setNStateInServer] ,
	[replacerId_fill] ,[replacerName_fill],[replacerId_lstFill] ,[replacerName_lstFill] ,[printTime] ,[wiId],[commitByDataWriter])
	 VALUES
	 (		  @rcid    --rcid
			   ,@rptId --<RtId, nvarchar(20),>
			   ,@DeptId--<fillDept, int,>
			   ,@DeptName--<fillDeptName, nvarchar(50),>
			   ,@UserId--<fillUser, int,>
			   ,@UserName--<fillUserName, nvarchar(50),>
			   ,getdate()--<fillDate, datetime,>
			   ,''--<rcDesc, nvarchar(2000),>
			   ,1--<state, smallint,>
			   ,1--<lstFiller, int,>
			   ,@UserName--<lstFillerName, nvarchar(50),>
			   ,getdate()--<lstFillDate, datetime,>
			   ,0--<backUpdate, smallint,>
			   ,0--<openState, smallint,>
			   ,NULL--<openBy, int,>
			   ,NULL--<openByName, nvarchar(50),>
			   ,NULL--<OpenBySesId, nvarchar(20),>
			   ,0--<lockState, smallint,>
			   ,0--<lockInServer, smallint,>
			   ,0--<noticeState, smallint,>
			   ,0--<setNStateInServer, smallint,>
			   ,0--<replacerId_fill, int,>
			   ,''--<replacerName_fill, nvarchar(50),>
			   ,0--<replacerId_lstFill, int,>
			   ,''--<replacerName_lstFill, nvarchar(50),>
			   ,NULL--<printTime, datetime,>
			   ,''--<wiId, nvarchar(20),>
			   ,0--<commitByDataWriter, smallint,>
	)

	end
```

### 相关附件下载
<a href="./3.9.zip" download>附件下载</a>

### 本节贡献者
*@Meteor*  
