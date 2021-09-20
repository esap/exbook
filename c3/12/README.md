# 3.12 SQL Server字符串聚合解决方法（CLR）

> @Meteor: 让字符串也能愉快的Sum起来  

引子：本文我只是实现方法的搬运工，三种方法均非原创，都是google上摘录。

SQL的聚合函数，例如SUM，AVG等主要针对数值，有时候不得不对字符串也做个按分组拼接。字符串也能SUM一下吗？

### 开发环境：SQL Server2008 R2
写个综合视图，遇到个情况，需要对字符串进行聚合统计，简化如下：

| 任务号  | 提交人  | 完工数  | 周转车号 |
| ---- | ---- | ---- | ---- |
| X01  | 张三   | 300  | V001 |
| X01  | 李四   | 200  | V002 |
| X02  | 王五   | 600  | V003 |
| X02  | 马六   | 400  | V004 |
| X02  | 赵七   | 100  | V005 |

目的是:需要列出统计任务的完成信息如下：

| 任务号  | 提交人      | 完工数  | 周转车号           |
| ---- | -------- | ---- | -------------- |
| X01  | 张三,李四    | 500  | V001,V002      |
| X02  | 王五,马六，赵七 | 1100 | V003,V004,V005 |

完工数量可以直接sum 后 group by,但是提交人 和 周转车 字符串字段就很麻烦了。google了下，有以下三种办法：

* ** 自定义聚合函数 **   [如何在sql server的group by语句中聚合字符串字段](https://zhidao.baidu.com/question/1431474448110090539.html)
  这种方法的思路就是用sql自定义个function，聚合的时候调用。这个办法最大的问题就是在函数中需要把要调用的表名写死，像上面这个需求，就要定义两个函数，一个是对提交人的聚合，一个是对周转车的聚合，而且这里的识别id只有一个，就是任务id（这个是简化需求），我的实际需求是要对任务ID+工序ID作为子件的，这样的函数条件也不好扩展。--所以放弃这个办法。
* ** 用stuff和for xml path子查询 **  [SQL SERVER 2005 中使用for xml path('')和stuff合并显示多行数据到一行中  ](http://lvmylove.blog.163.com/blog/static/207215172201511233315392/)
  这个方法也可行，但是问题也和1一样，要大段大段的写SQL子查询，而且无法复用，多的话实在受不了。
* ** 目前找到的以为最好的方法：配合c#自定义聚合函数  ** [源出处：C#实现SQL Server2005的扩展聚合函数](http://www.cnblogs.com/blues_/archive/2010/03/19/1690047.html)
  该方法实现后，调用的SQL就是：

```SQL
SELECT taskid,SUM(qty),
dbo.StrJoin(workerName,',') as workers, dbo.StrJoin(cartNo,',') as Carts 
FROM taskExecs  GROUP BY taskid
```
是不是很简单？而且以后出现类似的拼接字符串聚合就都直接调用就好了，一副一劳永逸的姿态。
我对原文的方法做了一些小调整和改变，具体实现如下：
1. Visual Studio 2015，新建个项目--》模板选SQL Server 数据库项目，命名项目sqlUtil
2. 新建项--》 SQL CLR c#  ==>SQL CLR c# 聚合  ，是个类，命名StrJoin.cs

```c#
using System;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using Microsoft.SqlServer.Server;
using System.Text;、
[Serializable]
[SqlUserDefinedAggregate(
  Format.UserDefined, //use custom serialization to serialize the intermediate result
  IsInvariantToNulls = true, //optimizer property
  IsInvariantToDuplicates = false, //optimizer property
  IsInvariantToOrder = false, //optimizer property
  MaxByteSize = 8000), //maximum size in bytes of persisted value
]
public struct StrJoin : IBinarySerialize
{
  private StringBuilder sbIntermediate;
  public void Init()
  {
    sbIntermediate = new StringBuilder();
  }
  public void Accumulate(SqlString Value,SqlString contChar)
  {
    if (Value == null || Value.ToString().ToLower().Equals("null"))
    {
      return;
    }
    else
    {
      sbIntermediate.Append(Value).Append(contChar);
    }
  }
  public void Merge(StrJoin Group)
  {
    sbIntermediate.Append(Group.sbIntermediate);
  }
  public SqlString Terminate()
  {
    string output = String.Empty;
    if (sbIntermediate != null && sbIntermediate.Length>0)
    {
      output = sbIntermediate.ToString(0, sbIntermediate.Length - 1);
    }
    return new SqlString(output);
  }
  // This is a place-holder member field
  #region IBinarySerialize Members
  public void Read(System.IO.BinaryReader r)
  {
    sbIntermediate = new StringBuilder(r.ReadString());
  }
  public void Write(System.IO.BinaryWriter w)
  {
    w.Write(this.sbIntermediate.ToString());
  }
  #endregion
}
```

说明：看上去一脸蒙逼很复杂的样子，其实以上函数有效的部分很简单，重点部分就是
1. 在Accumulate函数中：传入参数，把字符串拼起来。 
2. 在Terminate函数中： 去掉最后一个连接符并输出。
   主要看这两个动作，就知道了。
3. 实在不晓得.net就跳过这段，直接用附件中的sqlUtil.dll就好了。跳到下一步。
   在sqlserver中执行如下：

```SQL
--打开SQLSERVER的CLR功能
EXEC sp_configure 'clr enabled', 1
RECONFIGURE WITH OVERRIDE
GO
--注册DLL
CREATE ASSEMBLY sqlUtil FROM 'C:\sqlUtil.dll'      --生成的DLL路径
GO
--注册函数
CREATE AGGREGATE [dbo].[StrJoin] (@Value [nvarchar](MAX), 
  @contChar [nvarchar](2))
  RETURNS [nvarchar](MAX)
  EXTERNAL NAME [sqlUtil].[StrJoin]
```

这样后，就可以愉快的使用了。

如果要更新dll，需要先drop，在create

顺序是  删除引用的函数-->删除dll

    DROP AGGREGATE StrJoin
    DROP  ASSEMBLY sqlUtil

***PS:在这个过程遇到个纠结的问题，就是虚拟机和远程机之间复制文件的时候，居然会有问题，导致一个更新的dll一直是旧版本，而我却以为代码有错。。。。最后用.Net Refector去看dll的函数，才惊觉这个问题，吐血中.... 最后还是用共享传的文件。***

***PS2：据说MYSQL和Oracle其实都有现成的group_contact 和 wm_concat，所以到了SQLSERVER2012,据说也支持了字符串聚合的函数。但是在使用2012之前，等于是用第三种方法提前体验了而已。***

### 相关附件下载
[附件dll下载](c3/12/sqlUtil.dll ':ignore')

### 本节贡献者
*@Meteor*  
