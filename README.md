# LKUP_ConvertTextToDateSmall
Function to convert dates to text

One of the major issues with loading flat files, especially CSV files, are date conversion issues. Usually these occur from a region of the world with a different date format (or date format left at default USA format but everyone else in Europe is on something else) or from someone making the schoolboy error of opening the file in Microsoft Excel.

This function sits in a database called Lookups in my example. I use that database to store useful code that I reuse all over the place. The function itself is a table valued function so it should go parallel and scale relatively well.

An example of its use:

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
