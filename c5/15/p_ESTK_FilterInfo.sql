/****** Object:  StoredProcedure [dbo].[p_ESTK_FilterInfo]    Script Date: 06/25/2018 17:25:36 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE PROCEDURE [dbo].[p_ESTK_FilterInfo] (@filterTableName AS VARCHAR(100))
AS
	SELECT SheetId,RttId,DataRng,fldName,FldAlias,style FROM dbo.ES_v_Rtfs
	WHERE dtName=@filterTableName 
GO