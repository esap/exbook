# 工作流相关系统表

```sql
select * from JU_post --角色与部门结合
select * from JU_UserPost--userid与postid
select * from JU_TemplateReadRight
select * from JU_UserPost --用户角色表
JU_TemplateTableField--模板表列名。
JU_TemplateTable --模板表名。
select * from dbo.JU_Workflow  --流程名称
select * from dbo.JU_WorkflowActivity -- 流程活动
select * from dbo.JU_WorkflowDirection --流程放向
select * from dbo.JU_WorkflowDirectionDeliver --流程流向传递
select * from  dbo.JU_WorkflowInstance --流程实例
select * from dbo.JU_WorkflowTask --流程任务
select * from dbo.JU_WorkflowTaskLog --流程任务日志
```