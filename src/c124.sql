--高低配刷124版本专用，其它版本自己测试完善，by woylin 2016.3.24
declare @ver varchar(10)
set @ver='12.0.46'
--set @ver='9.4.124'
update ESModel..ES_HomeInfo	set Version=@ver

update ESSystem..ES_SysInfo	set Version=@ver

update esupg..UpdateInfo set Version=@ver

--刷应用版本，这里以esapp1和esap两个应用做例子
update ESApp1..ES_HomeInfo set Version=@ver
--update ESAp..ES_HomeInfo set Version=@ver

--补视图,解决124系统管理台点击应用报错
USE [ESSystem]
go
create view [dbo].[ES_v_Application]
as
select
a.AppId,a.AppName,a.Alias,a.Db,a.DbPath,a.CreTime
,case c.useNFS when 0 then 0 else a.NFSEnable end NFSEnable
,case c.useNFS when 0 then 0 else a.NFSMenuEnable end NFSMenuEnable
,case c.useNFS when 0 then 0 else a.saveFileInNFS end saveFileInNFS
,a.multiPwd,a.maxUn,
0 as dftFolder
from ES_Application a,ES_SysInfo c
GO

--补系统表，解决124网盘报错
USE [ESSystem]
GO
/****** Object:  Table [dbo].[ES_NFSFolder]    Script Date: 03/25/2016 10:39:29 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
--ES_NFSFolder
CREATE TABLE [dbo].[ES_NFSFolder](
	[FolderId] [int] NOT NULL,
	[isHostDir] [smallint] NOT NULL,
	[ServerId] [int] NOT NULL,
	[FolderName] [nvarchar](50) NOT NULL,
	[FolderDesc] [nvarchar](500) NULL,
	[Path] [nvarchar](420) NOT NULL,
	[AppId] [int] NULL,
	[useFor] [smallint] NOT NULL,
	[dftAttSpace] [smallint] NOT NULL,
	[hostDirId] [int] NOT NULL,
PRIMARY KEY CLUSTERED 
(
	[FolderId] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[ES_NFSFolder] ADD  DEFAULT ((0)) FOR [isHostDir]
GO
ALTER TABLE [dbo].[ES_NFSFolder] ADD  DEFAULT ((0)) FOR [useFor]
GO
ALTER TABLE [dbo].[ES_NFSFolder] ADD  DEFAULT ((0)) FOR [dftAttSpace]
GO
ALTER TABLE [dbo].[ES_NFSFolder] ADD  DEFAULT ((0)) FOR [hostDirId]
GO
--ES_NFSServer
CREATE TABLE [dbo].[ES_NFSServer](
	[ServerId] [int] NOT NULL,
	[ServerName] [nvarchar](50) NOT NULL,
	[Ip] [nvarchar](100) NOT NULL,
	[Port] [int] NOT NULL,
	[serverDesc] [nvarchar](500) NULL,
	[CreTime] [datetime] NULL,
PRIMARY KEY CLUSTERED 
(
	[ServerId] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[ES_NFSServer] ADD  DEFAULT (getdate()) FOR [CreTime]
GO
USE [ESModel]
--UserOption补丁
--ALTER TABLE [dbo].[ES_UserOption] ADD DftWin smallint default 1 not null 
GO
--开启定位，用于设计时
update ESApp1..ES_DataDomain set LocType=0 where DomainName='定位'
--激活定位，设计完后应用时再执行此行
--update ESApp1..ES_DataDomain set LocType=2 where DomainName='定位'

--开启二维码，注意先定义模型
update ESApp1..ES_DataDomain set Is2dBarcode=0 where DomainName='二维码'
