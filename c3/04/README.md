# 3.4 查询编辑中的单据的视图
```sql
select
	a.openbyname 使用者名称,
	b.RtName 模板名称,
	a.lstFillDate 打开时间
from
	ES_RepCase a
inner join
	ES_Tmp b
on
	a.RtId=b.RtId
where
	a.openState='1'
```

### 本节贡献者
*@毛毛*  
