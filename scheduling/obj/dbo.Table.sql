CREATE TABLE [dbo].[Table]
(
	[專案編號] NVARCHAR(50) NOT NULL PRIMARY KEY, 
    [採樣日期起] DATE NULL, 
    [委託單報告日期] DATE NULL, 
    [天數] INT NULL, 
    [檢測項目/設備名稱] NCHAR(10) NULL, 
    [分析方法] NVARCHAR(50) NULL, 
    [數量] INT NULL, 
    [課別] NCHAR(10) NULL, 
    [案件負責人] NCHAR(10) NULL
)
