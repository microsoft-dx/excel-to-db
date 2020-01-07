-- =======================================================
-- Create Stored Procedure Template for Azure SQL Database
-- =======================================================
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:      Xander
-- Create Date: 2020-01-07
-- Description: FromExcelToTable
-- =============================================
CREATE OR ALTER PROCEDURE FromExcelToTable (
	@tableName NVARCHAR(MAX)
	,@columns NVARCHAR(MAX)
	,@row NVARCHAR(MAX)
	)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON

	-- Insert statements for procedure here
	DECLARE @CUSTOM_SQL NVARCHAR(MAX) = 'INSERT INTO ' + @tableName + '(' + @columns + ')' + ' VALUES (' + @row + ')';

	PRINT (@CUSTOM_SQL);

	EXEC (@CUSTOM_SQL);
END
GO

