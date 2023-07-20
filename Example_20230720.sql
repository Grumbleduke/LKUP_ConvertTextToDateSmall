USE Lookups
GO
CREATE TABLE #t(DateAsText varchar(50) NOT NULL)
GO
INSERT INTO #t(DateAsText) VALUES('20211201')
INSERT INTO #t(DateAsText) VALUES('2021-12-01 00:00:00')
INSERT INTO #t(DateAsText) VALUES('01-Dec-2021')
INSERT INTO #t(DateAsText) VALUES('01/12/2021')
INSERT INTO #t(DateAsText) VALUES('01-12-2021')
GO
SELECT *,CONVERT(smalldatetime,'20211201',112) as DateSmall
FROM #t as t
	CROSS APPLY Lookups.dbo.tFnConvertStringToSmallDateTime(DateAsText) as d
GO
DROP TABLE #t
GO