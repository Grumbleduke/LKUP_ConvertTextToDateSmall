USE [Lookups]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE FUNCTION [dbo].[tFnConvertStringToSmallDateTime] --Function name

(@DateAsString varchar(30)) --Input variable specification

RETURNS TABLE --Output variable specification

AS RETURN
(

SELECT
/*-- 4 columns: 
StringAsDate					smalldatetime
AssumedDateFormatCode			int
DateConversionAccuracy			float
NumberOfPossibleTargetDateCodes tinyint
PossibleFormats					varchar(50)
*/
 CASE
-- Inconclusive stringmax or mix of dates
	WHEN LEN(@DateAsString)=26 AND @DateAsString LIKE '[a-Z][a-Z][a-Z] [0-3][0-9] [1,2][0-9][0-9][0-9] [0,1][0-9]:[0-5][0-9]:[0-5][0-9]:[0-9][0-9][0-9][A,P]M'	THEN CONVERT(smalldatetime,@DateAsString,109)--109 = mon dd yyyy hh:mi:ss:mmmAM (or PM)
	WHEN LEN(@DateAsString)=24 AND @DateAsString LIKE '[a-Z][a-Z][a-Z] [0-3][0-9] [0-9][0-9] [0,1][0-9]:[0-5][0-9]:[0-5][0-9]:[0-9][0-9][0-9][A,P]M'			THEN CONVERT(smalldatetime,@DateAsString,  9)--  9 = mon dd yy hh:mi:ss:mmmAM (or PM)
	WHEN LEN(@DateAsString)=24 AND @DateAsString LIKE '[1,2][0-9][0-9][0-9]-[0,1][0-9]-[0-3][0-9]T[0-2][0-9]:[0-5][0-9]:[0-5][0-9].[0-9][0-9][0-9]Z'			THEN CONVERT(smalldatetime,@DateAsString,127)--127 = yyyy-mm-ddThh:mi:ss.mmmZ
	WHEN LEN(@DateAsString)=24 AND @DateAsString LIKE '[0-3][0-9] [a-Z][a-Z][a-Z] [1,2][0-9][0-9][0-9] [0-9][0-9]:[0-9][0-9]:[0-9][0-9]:[0-9][0-9][0-9]'		THEN CONVERT(smalldatetime,@DateAsString,113)--113 = dd mon yyyy hh:mi:ss:mmm(24h)
	WHEN LEN(@DateAsString)=20 AND @DateAsString LIKE '[0-3][0-9] [a-Z][a-Z][a-Z] [1,2][0-9][0-9][0-9] [0-9][0-9]:[0-9][0-9]:[0-9][0-9]'						THEN CONVERT(smalldatetime,@DateAsString,113)--113 = dd mon yyyy hh:mi:ss(24h)
	WHEN LEN(@DateAsString)=22 AND @DateAsString LIKE '[0-3][0-9] [a-Z][a-Z][a-Z] [0-9][0-9] [0-9][0-9]:[0-9][0-9]:[0-9][0-9]:[0-9][0-9][0-9]'					THEN CONVERT(smalldatetime,@DateAsString, 13)-- 13 = dd mon yy hh:mi:ss:mmm(24h)
	WHEN LEN(@DateAsString)=23 AND @DateAsString LIKE '[1,2][0-9][0-9][0-9]-[0,1][0-9]-[0-3][0-9] [0-2][0-9]:[0-5][0-9]:[0-5][0-9].[0-9][0-9][0-9]'				THEN CONVERT(smalldatetime,@DateAsString,121)--121 = yyyy-mm-dd hh:mi:ss.mmm(24h)
	WHEN LEN(@DateAsString)=23 AND @DateAsString LIKE '[1,2][0-9][0-9][0-9]-[0,1][0-9]-[0-3][0-9]T[0-2][0-9]:[0-5][0-9]:[0-5][0-9].[0-9][0-9][0-9]'				THEN CONVERT(smalldatetime,@DateAsString,126)--126 = yyyy-mm-ddThh:mi:ss.mmm (no spaces)
	WHEN LEN(@DateAsString)=19 AND @DateAsString LIKE '[0-2][0-9][0-9][0-9]-[0,1][0-9]-[0-3][0-9] [0-2][0-9]:[0-5][0-9]:[0-5][0-9]'								THEN CONVERT(smalldatetime,@DateAsString,120)--120 = yyyy-mm-dd hh:mi:ss(24h)
	WHEN LEN(@DateAsString)=19 AND @DateAsString LIKE '[A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y] [0-3][0-9] [0-2][0-9][0-9][0-9] [0,1][0-9]:[0-5][0-9][A,P]M' THEN CONVERT(smalldatetime,@DateAsString,100)--100 = mon dd yyyy hh:miAM (or PM)
	WHEN LEN(@DateAsString)=17 AND @DateAsString LIKE '[A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y] [0-3][0-9] [0-9][0-9] [0,1][0-9]:[0-5][0-9][A,P]M'	THEN CONVERT(smalldatetime,@DateAsString,  0)--  0 = mon dd yy hh:miAM (or PM)
	WHEN LEN(@DateAsString)=12 AND @DateAsString LIKE '[0-2][0-9]:[0-5][0-9]:[0-5][0-9]:[0-9][0-9][0-9]'														THEN CONVERT(smalldatetime,@DateAsString,114)--114 = hh:mi:ss:mmm(24h)
	WHEN LEN(@DateAsString)=10 AND SUBSTRING(@DateAsString,4,2) < '13'
								AND @DateAsString LIKE '[0-3][0-9]/[0,1][0-9]/[0-2][0-9][0-9][0-9]'																THEN CONVERT(smalldatetime,@DateAsString,103)--103 = dd/mm/yyyy
	WHEN LEN(@DateAsString)=16 AND @DateAsString LIKE '[0-3][0-9]/[0,1][0-9]/[0-2][0-9][0-9][0-9] [0-2][0-9]:[0-5][0-9]'										THEN CONVERT(smalldatetime,LEFT(@DateAsString,10),103)+CONVERT(smalldatetime,RIGHT(@DateAsString,5),108)
	WHEN LEN(@DateAsString)= 8 AND SUBSTRING(@DateAsString,1,2) < '32' AND SUBSTRING(@DateAsString,4,2) < '13' 
								AND @DateAsString LIKE '[0-3][0-9]/[0,1][0-9]/[0-9][0-9]'																		THEN CONVERT(smalldatetime,@DateAsString,  3)--  3 = dd/mm/yy
	WHEN LEN(@DateAsString)=10 AND SUBSTRING(@DateAsString,1,2) < '13'
								AND @DateAsString LIKE '[0,1][0-9]/[0-3][0-9]/[0-2][0-9][0-9][0-9]'																THEN CONVERT(smalldatetime,@DateAsString,101)--101 = mm/dd/yyyy
	WHEN LEN(@DateAsString)=20 AND SUBSTRING(@DateAsString,1,2) < '13'
								AND @DateAsString LIKE '[0,1][0-9]/[0-3][0-9]/[0-9][0-9] [ ,0,1][0-9]:[0-5][0-9]:[0-5][0-9] [A,P]M'								THEN CONVERT(smalldatetime,@DateAsString,22)  -- 22 = mm/dd/yy hh:mi:ss AM (or PM)
	WHEN LEN(@DateAsString)=10 AND @DateAsString LIKE '[0-2][0-9][0-9][0-9]/[0,1][0-9]/[0-3][0-9]'																THEN CONVERT(smalldatetime,@DateAsString,111)--111 = yyyy/mm/dd
	WHEN LEN(@DateAsString)= 8 AND SUBSTRING(@DateAsString,4,2) < '13' AND SUBSTRING(@DateAsString,6,2) < '32' 
								AND @DateAsString LIKE '[0-9][0-9]/[0,1][0-9]/[0-3][0-9]'																		THEN CONVERT(smalldatetime,@DateAsString, 11)-- 11 = yy/mm/dd
	WHEN LEN(@DateAsString)= 8 AND SUBSTRING(@DateAsString,1,2) < '13' AND SUBSTRING(@DateAsString,4,2) < '32' 
								AND @DateAsString LIKE '[0,1][0-9]/[0-3][0-9]/[0-9][0-9]'																		THEN CONVERT(smalldatetime,@DateAsString,  1)--  1 = mm/dd/yy
	WHEN LEN(@DateAsString)=10 AND @DateAsString LIKE '[A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y] [0-3][0-9], [0-9][0-9]'								THEN CONVERT(smalldatetime,@DateAsString,  7)--  7 = Mon dd, yy
	WHEN LEN(@DateAsString)=12 AND @DateAsString LIKE '[A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y] [0-3][0-9], [0-2][0-9][0-9][0-9]'					THEN CONVERT(smalldatetime,@DateAsString,107)--107 = Mon dd, yyyy
	WHEN LEN(@DateAsString)= 9 AND @DateAsString LIKE '[0-3][0-9][ ,-][A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y][ ,-][0-9][0-9]'						THEN CONVERT(smalldatetime,@DateAsString,  6)--  6 = dd mon yy
	WHEN LEN(@DateAsString)=11 AND @DateAsString LIKE '[0-3][0-9][ ,-][A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y][ ,-][0-2][0-9][0-9][0-9]'				THEN CONVERT(smalldatetime,@DateAsString,106)--106 = dd mon yyyy
	WHEN LEN(@DateAsString)= 8 AND @DateAsString LIKE '[0-9][0-9].[0,1][0-9].[0-3][0-9]'																		THEN CONVERT(smalldatetime,@DateAsString,  2)--  2 = yy.mm.dd
	WHEN LEN(@DateAsString)=10 AND @DateAsString LIKE '[0-2][0-9][0-9][0-9].[0,1][0-9].[0-3][0-9]'																THEN CONVERT(smalldatetime,@DateAsString,102)--102 = yyyy.mm.dd
	WHEN LEN(@DateAsString)= 8 AND @DateAsString LIKE '[0-3][0-9].[0,1][0-9].[0-9][0-9]'																		THEN CONVERT(smalldatetime,@DateAsString,  4)--  4 = dd.mm.yy
	WHEN LEN(@DateAsString)=10 AND @DateAsString LIKE '[0-3][0-9].[0,1][0-9].[0-2][0-9][0-9][0-9]'																THEN CONVERT(smalldatetime,@DateAsString,104)--104 = dd.mm.yyyy
	WHEN LEN(@DateAsString)= 8 AND @DateAsString LIKE '[0-3][0-9]-[0,1][0-9]-[0-9][0-9]'																		THEN CONVERT(smalldatetime,@DateAsString,  5)--  5 = dd-mm-yy
	WHEN LEN(@DateAsString)=10 AND @DateAsString LIKE '[0-3][0-9]-[0,1][0-9]-[0-2][0-9][0-9][0-9]'																THEN CONVERT(smalldatetime,@DateAsString,105)--105 = dd-mm-yyyy
	WHEN LEN(@DateAsString)= 8 AND @DateAsString LIKE '[0-2][0-9]:[0-5][0-9]:[0-5][0-9]'																		THEN CONVERT(smalldatetime,@DateAsString,108)--108 = hh:mi:ss
	WHEN LEN(@DateAsString)= 7 AND @DateAsString LIKE '[0-9]:[0-5][0-9]:[0-5][0-9]'																				THEN CONVERT(smalldatetime,@DateAsString,108)--108 = hh:mi:ss - leading zero truncated
	WHEN LEN(@DateAsString)= 8 AND @DateAsString LIKE '[1,2][0-9][0-9][0-9][0,1][0-9][0-3][0-9]'																THEN CONVERT(smalldatetime,@DateAsString,112)--112 = yyyymmdd
	WHEN LEN(@DateAsString)= 8 AND @DateAsString LIKE '[0,1][0-9]-[0-3][0-9]-[0-9][0-9]'																		THEN CONVERT(smalldatetime,@DateAsString, 10)-- 10 = mm-dd-yy
	WHEN LEN(@DateAsString)= 6 AND @DateAsString LIKE '[0-9][0-9][0,1][0-9][0-3][0-9]'																			THEN CONVERT(smalldatetime,@DateAsString, 12)-- 12 = yymmdd
	WHEN LEN(@DateAsString)=26 AND @DateAsString LIKE '[0-3][0-9] [a-Z][a-Z][a-Z] [1,2][0-9][0-9][0-9] [0,1][0-9]:[0-5][0-9]:[0-5][0-9]:[0-9][0-9][0-9][A,P]M'	THEN CONVERT(smalldatetime,@DateAsString/*,130*/)--130 = dd mon yyyy hh:mi:ss:mmmAM -- Stating 130 causes error!
	--WHEN LEN(@DateAsString)=23 AND @DateAsString LIKE '[0-3][0-9]/[0,1][0-9]/[0-9][0-9] [0,1][0-9]:[0-5][0-9]:[0-5][0-9]:[0-9][0-9][0-9][A,P]M'				THEN CONVERT(smalldatetime,@DateAsString/*,131*/)--131 = dd/mm/yy hh:mi:ss:mmmAM -- Bombs out with error!
	WHEN LEN(@DateAsString)=21 AND @DateAsString LIKE '[0-9][0-9]-[0,1][0-9]-[0-3][0-9] [0-2][0-9]:[0-5][0-9]:[0-5][0-9].[0-9][0-9][0-9]'						THEN CONVERT(smalldatetime,@DateAsString/*, 21*/)-- 21 = yy-mm-dd hh:mi:ss.mmm(24h) -- Stating 21 causes error!
	WHEN LEN(@DateAsString)=17 AND @DateAsString LIKE '[0-9][0-9]-[0,1][0-9]-[0-3][0-9] [0-2][0-9]:[0-5][0-9]:[0-5][0-9]'										THEN CONVERT(smalldatetime,@DateAsString/*, 20*/)-- 20 = yy-mm-dd hh:mi:ss(24h) -- Stating 20 causes error!
	WHEN LEN(@DateAsString)>19 AND @DateAsString LIKE '[0-2][0-9][0-9][0-9]-[0,1][0-9]-[0-3][0-9] [0-2][0-9]:[0-5][0-9]:[0-5][0-9].0000000'						THEN CONVERT(smalldatetime,LEFT(@DateAsString,19),120)--120 = yyyy-mm-dd hh:mi:ss(24h)
--EXCEL HELL FOLLOWS
	WHEN @DateAsString LIKE '[0-9]/[0-9]/[1,2][0-9][0-9][0-9] [0-9]:[0-5][0-9]' THEN CONVERT(smalldatetime,SUBSTRING(@DateAsString,5,4)+'-0'+SUBSTRING(@DateAsString,1,1)+'-0'+SUBSTRING(@DateAsString,3,1)+' '+'0'+SUBSTRING(@DateAsString,10,1)+RIGHT(@DateAsString,3)+':00',120)
	WHEN @DateAsString LIKE '[0-9]/[0-9]/[1,2][0-9][0-9][0-9] [0-2][0-9]:[0-5][0-9]' THEN CONVERT(smalldatetime,SUBSTRING(@DateAsString,5,4)+'-0'+SUBSTRING(@DateAsString,1,1)+'-0'+SUBSTRING(@DateAsString,3,1)+' '+SUBSTRING(@DateAsString,10,2)+RIGHT(@DateAsString,3)+':00',120)
	WHEN @DateAsString LIKE '[0-9]/[0-3][0-9]/[1,2][0-9][0-9][0-9] [0-9]:[0-5][0-9]' THEN CONVERT(smalldatetime,SUBSTRING(@DateAsString,6,4)+'-'+SUBSTRING(@DateAsString,1,1)+'-0'+SUBSTRING(@DateAsString,3,2)+' '+'0'+SUBSTRING(@DateAsString,11,1)+RIGHT(@DateAsString,3)+':00',120)
	WHEN @DateAsString LIKE '[0-9]/[0-3][0-9]/[1,2][0-9][0-9][0-9] [0-5][0-9]:[0-5][0-9]'	THEN CONVERT(smalldatetime,SUBSTRING(@DateAsString,6,4)+'-0'+SUBSTRING(@DateAsString,1,1)+'-'+SUBSTRING(@DateAsString,3,2)+' '+SUBSTRING(@DateAsString,11,2)+RIGHT(@DateAsString,3)+':00',120)
	WHEN @DateAsString IS NOT NULL AND ISDATE(@DateAsString)=0 AND @DateAsString < 'a' AND @DateAsString NOT LIKE '%-%' AND @DateAsString LIKE '%.%' THEN CONVERT(varchar(20),CAST(CAST(@DateAsString as float) - 2.00 as smalldatetime),113)
	WHEN @DateAsString  LIKE '[0-9][0-9][0-9][0-9][0-9]' THEN DATEADD(DAY,CAST(@DateAsString as int),'30-Dec-1899') 
	WHEN @DateAsString LIKE '[0-9][ ,-,/][0-9][ ,-,/][1,2][0-9][0-9][0-9]' THEN CONVERT(smalldatetime,SUBSTRING(@DateAsString,5,4)+'-0'+SUBSTRING(@DateAsString,3,1)+'-0'+SUBSTRING(@DateAsString,1,1)+' 00:00:00',120)
	ELSE CAST(NULL as datetime)
	END as StringAsDate

,CAST(CASE
-- Inconclusive stringmax or mix of dates
	WHEN LEN(@DateAsString)=26 AND @DateAsString LIKE '[a-Z][a-Z][a-Z] [0-3][0-9] [1,2][0-9][0-9][0-9] [0,1][0-9]:[0-5][0-9]:[0-5][0-9]:[0-9][0-9][0-9][A,P]M'	THEN 109--109 = mon dd yyyy hh:mi:ss:mmmAM (or PM)
	WHEN LEN(@DateAsString)=24 AND @DateAsString LIKE '[a-Z][a-Z][a-Z] [0-3][0-9] [0-9][0-9] [0,1][0-9]:[0-5][0-9]:[0-5][0-9]:[0-9][0-9][0-9][A,P]M'			THEN   9--  9 = mon dd yy hh:mi:ss:mmmAM (or PM)
	WHEN LEN(@DateAsString)=24 AND @DateAsString LIKE '[1,2][0-9][0-9][0-9]-[0,1][0-9]-[0-3][0-9]T[0-2][0-9]:[0-5][0-9]:[0-5][0-9].[0-9][0-9][0-9]Z'			THEN 127--127 = yyyy-mm-ddThh:mi:ss.mmmZ
	WHEN LEN(@DateAsString)=24 AND @DateAsString LIKE '[0-3][0-9] [a-Z][a-Z][a-Z] [1,2][0-9][0-9][0-9] [0-9][0-9]:[0-9][0-9]:[0-9][0-9]:[0-9][0-9][0-9]'		THEN 113--113 = dd mon yyyy hh:mi:ss:mmm(24h)
	WHEN LEN(@DateAsString)=20 AND @DateAsString LIKE '[0-3][0-9] [a-Z][a-Z][a-Z] [1,2][0-9][0-9][0-9] [0-9][0-9]:[0-9][0-9]:[0-9][0-9]'						THEN 113--113 = dd mon yyyy hh:mi:ss:mmm(24h)
	WHEN LEN(@DateAsString)=22 AND @DateAsString LIKE '[0-3][0-9] [a-Z][a-Z][a-Z] [0-9][0-9] [0-9][0-9]:[0-9][0-9]:[0-9][0-9]:[0-9][0-9][0-9]'					THEN  13-- 13 = dd mon yy hh:mi:ss:mmm(24h)
	WHEN LEN(@DateAsString)=23 AND @DateAsString LIKE '[1,2][0-9][0-9][0-9]-[0,1][0-9]-[0-3][0-9] [0-2][0-9]:[0-5][0-9]:[0-5][0-9].[0-9][0-9][0-9]'				THEN 121--121 = yyyy-mm-dd hh:mi:ss.mmm(24h)
	WHEN LEN(@DateAsString)=23 AND @DateAsString LIKE '[1,2][0-9][0-9][0-9]-[0,1][0-9]-[0-3][0-9]T[0-2][0-9]:[0-5][0-9]:[0-5][0-9].[0-9][0-9][0-9]'				THEN 126--126 = yyyy-mm-ddThh:mi:ss.mmm (no spaces)
	WHEN LEN(@DateAsString)=19 AND @DateAsString LIKE '[0-2][0-9][0-9][0-9]-[0,1][0-9]-[0-3][0-9] [0-2][0-9]:[0-5][0-9]:[0-5][0-9]'								THEN 120--120 = yyyy-mm-dd hh:mi:ss(24h)
	WHEN LEN(@DateAsString)=19 AND @DateAsString LIKE '[A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y] [0-3][0-9] [0-2][0-9][0-9][0-9] [0,1][0-9]:[0-5][0-9][A,P]M' THEN 100--100 = mon dd yyyy hh:miAM (or PM)
	WHEN LEN(@DateAsString)=17 AND @DateAsString LIKE '[A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y] [0-3][0-9] [0-9][0-9] [0,1][0-9]:[0-5][0-9][A,P]M'	THEN   0--  0 = mon dd yy hh:miAM (or PM)
	WHEN LEN(@DateAsString)=12 AND @DateAsString LIKE '[0-2][0-9]:[0-5][0-9]:[0-5][0-9]:[0-9][0-9][0-9]'														THEN 114--114 = hh:mi:ss:mmm(24h)
	WHEN LEN(@DateAsString)=10 AND SUBSTRING(@DateAsString,4,2) < '13'
								AND @DateAsString LIKE '[0-3][0-9]/[0,1][0-9]/[0-2][0-9][0-9][0-9]'																THEN 103--103 = dd/mm/yyyy
	WHEN LEN(@DateAsString)=16 AND @DateAsString LIKE '[0-3][0-9]/[0,1][0-9]/[0-2][0-9][0-9][0-9] [0-2][0-9]:[0-5][0-9]'										THEN 211 --103+108
	WHEN LEN(@DateAsString)= 8 AND SUBSTRING(@DateAsString,1,2) < '32' AND SUBSTRING(@DateAsString,4,2) < '13' 
								AND @DateAsString LIKE '[0-3][0-9]/[0,1][0-9]/[0-9][0-9]'																		THEN   3--  3 = dd/mm/yy
	WHEN LEN(@DateAsString)=10 AND SUBSTRING(@DateAsString,1,2) < '13'
								AND @DateAsString LIKE '[0,1][0-9]/[0-3][0-9]/[0-2][0-9][0-9][0-9]'																THEN 101--101 = mm/dd/yyyy
	WHEN LEN(@DateAsString)=20 AND SUBSTRING(@DateAsString,1,2) < '13'
								AND @DateAsString LIKE '[0,1][0-9]/[0-3][0-9]/[0-9][0-9] [ ,0,1][0-9]:[0-5][0-9]:[0-5][0-9] [A,P]M'								THEN 22 -- 22 = mm/dd/yy hh:mi:ss AM (or PM)
	WHEN LEN(@DateAsString)=10 AND @DateAsString LIKE '[0-2][0-9][0-9][0-9]/[0,1][0-9]/[0-3][0-9]'																THEN 111--111 = yyyy/mm/dd
	WHEN LEN(@DateAsString)= 8 AND SUBSTRING(@DateAsString,4,2) < '13' AND SUBSTRING(@DateAsString,6,2) < '32' 
								AND @DateAsString LIKE '[0-9][0-9]/[0,1][0-9]/[0-3][0-9]'																		THEN  11-- 11 = yy/mm/dd
	WHEN LEN(@DateAsString)= 8 AND SUBSTRING(@DateAsString,1,2) < '13' AND SUBSTRING(@DateAsString,4,2) < '32' 
								AND @DateAsString LIKE '[0,1][0-9]/[0-3][0-9]/[0-9][0-9]'																		THEN   1--  1 = mm/dd/yy
	WHEN LEN(@DateAsString)=10 AND @DateAsString LIKE '[A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y] [0-3][0-9], [0-9][0-9]'								THEN   7--  7 = Mon dd, yy
	WHEN LEN(@DateAsString)=12 AND @DateAsString LIKE '[A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y] [0-3][0-9], [0-2][0-9][0-9][0-9]'					THEN 107--107 = Mon dd, yyyy
	WHEN LEN(@DateAsString)= 9 AND @DateAsString LIKE '[0-3][0-9][ ,-][A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y][ ,-][0-9][0-9]'						THEN   6--  6 = dd mon yy
	WHEN LEN(@DateAsString)=11 AND @DateAsString LIKE '[0-3][0-9][ ,-][A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y][ ,-][0-2][0-9][0-9][0-9]'				THEN 106--106 = dd mon yyyy
	WHEN LEN(@DateAsString)= 8 AND @DateAsString LIKE '[0-9][0-9].[0,1][0-9].[0-3][0-9]'																		THEN   2--  2 = yy.mm.dd
	WHEN LEN(@DateAsString)=10 AND @DateAsString LIKE '[0-2][0-9][0-9][0-9].[0,1][0-9].[0-3][0-9]'																THEN 102--102 = yyyy.mm.dd
	WHEN LEN(@DateAsString)= 8 AND @DateAsString LIKE '[0-3][0-9].[0,1][0-9].[0-9][0-9]'																		THEN   4--  4 = dd.mm.yy
	WHEN LEN(@DateAsString)=10 AND @DateAsString LIKE '[0-3][0-9].[0,1][0-9].[0-2][0-9][0-9][0-9]'																THEN 104--104 = dd.mm.yyyy
	WHEN LEN(@DateAsString)= 8 AND @DateAsString LIKE '[0-3][0-9]-[0,1][0-9]-[0-9][0-9]'																		THEN   5--  5 = dd-mm-yy
	WHEN LEN(@DateAsString)=10 AND @DateAsString LIKE '[0-3][0-9]-[0,1][0-9]-[0-2][0-9][0-9][0-9]'																THEN 105--105 = dd-mm-yyyy
	WHEN LEN(@DateAsString)= 8 AND @DateAsString LIKE '[0-2][0-9]:[0-5][0-9]:[0-5][0-9]'																		THEN 108--108 = hh:mi:ss
	WHEN LEN(@DateAsString)= 7 AND @DateAsString LIKE '[0-9]:[0-5][0-9]:[0-5][0-9]'																				THEN -108--108 = hh:mi:ss - leading zero truncated
	WHEN LEN(@DateAsString)= 8 AND @DateAsString LIKE '[1,2][0-9][0-9][0-9][0,1][0-9][0-3][0-9]'																THEN 112--112 = yyyymmdd
	WHEN LEN(@DateAsString)= 8 AND @DateAsString LIKE '[0,1][0-9]-[0-3][0-9]-[0-9][0-9]'																		THEN  10-- 10 = mm-dd-yy
	WHEN LEN(@DateAsString)= 6 AND @DateAsString LIKE '[0-9][0-9][0,1][0-9][0-3][0-9]'																			THEN  12-- 12 = yymmdd
	WHEN LEN(@DateAsString)=26 AND @DateAsString LIKE '[0-3][0-9] [a-Z][a-Z][a-Z] [1,2][0-9][0-9][0-9] [0,1][0-9]:[0-5][0-9]:[0-5][0-9]:[0-9][0-9][0-9][A,P]M'	THEN 130--130 = dd mon yyyy hh:mi:ss:mmmAM -- Stating 130 causes error!
	WHEN LEN(@DateAsString)=23 AND @DateAsString LIKE '[0-3][0-9]/[0,1][0-9]/[0-9][0-9] [0,1][0-9]:[0-5][0-9]:[0-5][0-9]:[0-9][0-9][0-9][A,P]M'					THEN 131--131 = dd/mm/yy hh:mi:ss:mmmAM -- Bombs out with error!
	WHEN LEN(@DateAsString)=21 AND @DateAsString LIKE '[0-9][0-9]-[0,1][0-9]-[0-3][0-9] [0-2][0-9]:[0-5][0-9]:[0-5][0-9].[0-9][0-9][0-9]'						THEN  21-- 21 = yy-mm-dd hh:mi:ss.mmm(24h) -- Stating 21 causes error!
	WHEN LEN(@DateAsString)=17 AND @DateAsString LIKE '[0-9][0-9]-[0,1][0-9]-[0-3][0-9] [0-2][0-9]:[0-5][0-9]:[0-5][0-9]'										THEN  20-- 20 = yy-mm-dd hh:mi:ss(24h) -- Stating 20 causes error!
	WHEN LEN(@DateAsString)>19 AND @DateAsString LIKE '[0-2][0-9][0-9][0-9]-[0,1][0-9]-[0-3][0-9] [0-2][0-9]:[0-5][0-9]:[0-5][0-9].0000000'						THEN 120--120 = yyyy-mm-dd hh:mi:ss(24h)
--	WHEN @DateAsString IS NOT NULL AND ISDATE(@DateAsString)=0 AND @DateAsString < 'a' AND @DateAsString NOT LIKE '%-%' AND @DateAsString LIKE '%.%' THEN CONVERT(varchar(20),CAST(CAST(@DateAsString as float) - 2.00 as smalldatetime),113)
--	WHEN @DateAsString  LIKE '[0-9][0-9][0-9][0-9][0-9]'																										THEN -1 
	--ELSE 999
	END as int) as AssumedDateFormatCode

,ROUND(CASE
		WHEN @DateAsString IS NULL 
			THEN NULL
		WHEN CAST(-- number of possibilities ...
			 CASE WHEN LEN(@DateAsString)=26 AND @DateAsString LIKE '[a-Z][a-Z][a-Z] [0-3][0-9] [1,2][0-9][0-9][0-9] [0,1][0-9]:[0-5][0-9]:[0-5][0-9]:[0-9][0-9][0-9][A,P][M]'	THEN 1 ELSE 0 END--109 = mon dd yyyy hh:mi:ss:mmmAM (or PM)
			+CASE WHEN LEN(@DateAsString)=24 AND @DateAsString LIKE '[a-Z][a-Z][a-Z] [0-3][0-9] [0-9][0-9] [0,1][0-9]:[0-5][0-9]:[0-5][0-9]:[0-9][0-9][0-9][A,P][M]'			THEN 1 ELSE 0 END--  9 = mon dd yy hh:mi:ss:mmmAM (or PM)
			+CASE WHEN LEN(@DateAsString)=24 AND @DateAsString LIKE '[1,2][0-9][0-9][0-9]-[0,1][0-9]-[0-3][0-9]T[0-2][0-9]:[0-5][0-9]:[0-5][0-9].[0-9][0-9][0-9]Z'				THEN 1 ELSE 0 END--127 = yyyy-mm-ddThh:mi:ss.mmmZ
			+CASE WHEN LEN(@DateAsString)=24 AND @DateAsString LIKE '[0-3][0-9] [a-Z][a-Z][a-Z] [1,2][0-9][0-9][0-9] [0-9][0-9]:[0-9][0-9]:[0-9][0-9]:[0-9][0-9][0-9]'			THEN 1 ELSE 0 END--113 = dd mon yyyy hh:mi:ss:mmm(24h)
			+CASE WHEN LEN(@DateAsString)=24 AND @DateAsString LIKE '[0-3][0-9] [a-Z][a-Z][a-Z] [1,2][0-9][0-9][0-9] [0-9][0-9]:[0-9][0-9]:[0-9][0-9]'							THEN 1 ELSE 0 END--113 = dd mon yyyy hh:mi:ss:mmm(24h)
			+CASE WHEN LEN(@DateAsString)=22 AND @DateAsString LIKE '[0-3][0-9] [a-Z][a-Z][a-Z] [0-9][0-9] [0-9][0-9]:[0-9][0-9]:[0-9][0-9]:[0-9][0-9][0-9]'					THEN 1 ELSE 0 END-- 13 = dd mon yy hh:mi:ss:mmm(24h)
			+CASE WHEN LEN(@DateAsString)=23 AND @DateAsString LIKE '[1,2][0-9][0-9][0-9]-[0,1][0-9]-[0-3][0-9] [0-2][0-9]:[0-5][0-9]:[0-5][0-9].[0-9][0-9][0-9]'				THEN 1 ELSE 0 END--121 = yyyy-mm-dd hh:mi:ss.mmm(24h)
			+CASE WHEN LEN(@DateAsString)=23 AND @DateAsString LIKE '[1,2][0-9][0-9][0-9]-[0,1][0-9]-[0-3][0-9]T[0-2][0-9]:[0-5][0-9]:[0-5][0-9].[0-9][0-9][0-9]'				THEN 1 ELSE 0 END--126 = yyyy-mm-ddThh:mi:ss.mmm (no spaces)
			+CASE WHEN LEN(@DateAsString)=19 AND @DateAsString LIKE '[0-2][0-9][0-9][0-9]-[0,1][0-9]-[0-3][0-9] [0-2][0-9]:[0-5][0-9]:[0-5][0-9]'								THEN 1 ELSE 0 END--120 = yyyy-mm-dd hh:mi:ss(24h)
			+CASE WHEN LEN(@DateAsString)=19 AND @DateAsString LIKE '[A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y] [0-3][0-9] [0-2][0-9][0-9][0-9] [0,1][0-9]:[0-5][0-9][A,P]M' THEN 1 ELSE 0 END--100 = mon dd yyyy hh:miAM (or PM)
			+CASE WHEN LEN(@DateAsString)=17 AND @DateAsString LIKE '[A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y] [0-3][0-9] [0-9][0-9] [0,1][0-9]:[0-5][0-9][A,P]M'		THEN 1 ELSE 0 END--  0 = mon dd yy hh:miAM (or PM)
			+CASE WHEN LEN(@DateAsString)=12 AND @DateAsString LIKE '[0-2][0-9]:[0-5][0-9]:[0-5][0-9]:[0-9][0-9][0-9]'															THEN 1 ELSE 0 END--114 = hh:mi:ss:mmm(24h)
			+CASE WHEN LEN(@DateAsString)=10 AND SUBSTRING(@DateAsString,4,2) < '13'
												AND @DateAsString LIKE '[0-3][0-9]/[0,1][0-9]/[0-2][0-9][0-9][0-9]'																THEN 1 ELSE 0 END--103 = dd/mm/yyyy
			+CASE WHEN LEN(@DateAsString)=16 AND @DateAsString LIKE '[0-3][0-9]/[0,1][0-9]/[0-2][0-9][0-9][0-9] [0-2][0-9]:[0-5][0-9]'											THEN 1 ELSE 0 END --103+108
			+CASE WHEN LEN(@DateAsString)= 8 AND SUBSTRING(@DateAsString,1,2) < '32' AND SUBSTRING(@DateAsString,4,2) < '13' 
											AND @DateAsString LIKE '[0-3][0-9]/[0,1][0-9]/[0-9][0-9]'																			THEN 1 ELSE 0 END--  3 = dd/mm/yy
			+CASE WHEN LEN(@DateAsString)=10 AND SUBSTRING(@DateAsString,1,2) < '13'
												AND @DateAsString LIKE '[0,1][0-9]/[0-3][0-9]/[0-2][0-9][0-9][0-9]'																THEN 1 ELSE 0 END--101 = mm/dd/yyyy
			+CASE WHEN LEN(@DateAsString)=20 AND SUBSTRING(@DateAsString,1,2) < '13'
												AND @DateAsString LIKE '[0,1][0-9]/[0-3][0-9]/[0-9][0-9] [ ,0,1][0-9]:[ ,0-5][0-9]:[0-5][0-9] [A,P]M'							THEN 1 ELSE 0 END-- 22 = mm/dd/yy hh:mi:ss AM (or PM)
			+CASE WHEN LEN(@DateAsString)=10 AND @DateAsString LIKE '[0-2][0-9][0-9][0-9]/[0,1][0-9]/[0-3][0-9]'																THEN 1 ELSE 0 END--111 = yyyy/mm/dd
			+CASE WHEN LEN(@DateAsString)= 8 AND SUBSTRING(@DateAsString,4,2) < '13' AND SUBSTRING(@DateAsString,6,2) < '32' 
											AND @DateAsString LIKE '[0-9][0-9]/[0,1][0-9]/[0-3][0-9]'																			THEN 1 ELSE 0 END-- 11 = yy/mm/dd
			+CASE WHEN LEN(@DateAsString)= 8 AND SUBSTRING(@DateAsString,1,2) < '13' AND SUBSTRING(@DateAsString,4,2) < '32' 
											AND @DateAsString LIKE '[0,1][0-9]/[0-3][0-9]/[0-9][0-9]'																			THEN 1 ELSE 0 END--  1 = mm/dd/yy
			+CASE WHEN LEN(@DateAsString)=10 AND @DateAsString LIKE '[A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y] [0-3][0-9], [0-9][0-9]'								THEN 1 ELSE 0 END--  7 = Mon dd, yy
			+CASE WHEN LEN(@DateAsString)=12 AND @DateAsString LIKE '[A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y] [0-3][0-9], [0-2][0-9][0-9][0-9]'						THEN 1 ELSE 0 END--107 = Mon dd, yyyy
			+CASE WHEN LEN(@DateAsString)= 9 AND @DateAsString LIKE '[0-3][0-9][ ,-][A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y][ ,-][0-9][0-9]'							THEN 1 ELSE 0 END--  6 = dd mon yy
			+CASE WHEN LEN(@DateAsString)=11 AND @DateAsString LIKE '[0-3][0-9][ ,-][A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y][ ,-][0-2][0-9][0-9][0-9]'				THEN 1 ELSE 0 END--106 = dd mon yyyy
			+CASE WHEN LEN(@DateAsString)= 8 AND @DateAsString LIKE '[0-9][0-9].[0,1][0-9].[0-3][0-9]'																			THEN 1 ELSE 0 END--  2 = yy.mm.dd
			+CASE WHEN LEN(@DateAsString)=10 AND @DateAsString LIKE '[0-2][0-9][0-9][0-9].[0,1][0-9].[0-3][0-9]'																THEN 1 ELSE 0 END--102 = yyyy.mm.dd
			+CASE WHEN LEN(@DateAsString)= 8 AND @DateAsString LIKE '[0-3][0-9].[0,1][0-9].[0-9][0-9]'																			THEN 1 ELSE 0 END--  4 = dd.mm.yy
			+CASE WHEN LEN(@DateAsString)=10 AND @DateAsString LIKE '[0-3][0-9].[0,1][0-9].[0-2][0-9][0-9][0-9]'																THEN 1 ELSE 0 END--104 = dd.mm.yyyy
			+CASE WHEN LEN(@DateAsString)= 8 AND @DateAsString LIKE '[0-3][0-9]-[0,1][0-9]-[0-9][0-9]'																			THEN 1 ELSE 0 END--  5 = dd-mm-yy
			+CASE WHEN LEN(@DateAsString)=10 AND @DateAsString LIKE '[0-3][0-9]-[0,1][0-9]-[0-2][0-9][0-9][0-9]'																THEN 1 ELSE 0 END--105 = dd-mm-yyyy
			+CASE WHEN LEN(@DateAsString)= 8 AND @DateAsString LIKE '[0-2][0-9]:[0-5][0-9]:[0-5][0-9]'																			THEN 1 ELSE 0 END--108 = hh:mi:ss
			+CASE WHEN LEN(@DateAsString)= 7 AND @DateAsString LIKE '[0-9]:[0-5][0-9]:[0-5][0-9]'																				THEN 1 ELSE 0 END--108 = hh:mi:ss - leading zero truncated
			+CASE WHEN LEN(@DateAsString)= 8 AND @DateAsString LIKE '[1,2][0-9][0-9][0-9][0,1][0-9][0-3][0-9]'																	THEN 1 ELSE 0 END--112 = yyyymmdd
			+CASE WHEN LEN(@DateAsString)= 8 AND @DateAsString LIKE '[0,1][0-9]-[0-3][0-9]-[0-9][0-9]'																			THEN 1 ELSE 0 END-- 10 = mm-dd-yy
			+CASE WHEN LEN(@DateAsString)= 6 AND @DateAsString LIKE '[0-9][0-9][0,1][0-9][0-3][0-9]'																			THEN 1 ELSE 0 END-- 12 = yymmdd
			+CASE WHEN LEN(@DateAsString)=26 AND @DateAsString LIKE '[0-3][0-9] [a-Z][a-Z][a-Z] [1,2][0-9][0-9][0-9] [0,1][0-9]:[0-5][0-9]:[0-5][0-9]:[0-9][0-9][0-9][A,P][M]'	THEN 1 ELSE 0 END--130 = dd mon yyyy hh:mi:ss:mmmAM -- Stating 130 causes error!
			--+CASE WHEN LEN(@DateAsString)=23 AND @DateAsString LIKE '[0-3][0-9]/[0,1][0-9]/[0-9][0-9] [0,1][0-9]:[0-5][0-9]:[0-5][0-9]:[0-9][0-9][0-9][A,P]M'					THEN 1  ELSE 0 END--131 = dd/mm/yy hh:mi:ss:mmmAM -- Bombs out with error!
			+CASE WHEN LEN(@DateAsString)=21 AND @DateAsString LIKE '[0-9][0-9]-[0,1][0-9]-[0-3][0-9] [0-2][0-9]:[0-5][0-9]:[0-5][0-9].[0-9][0-9][0-9]'							THEN 1 ELSE 0 END-- 21 = yy-mm-dd hh:mi:ss.mmm(24h) -- Stating 21 causes error!
			+CASE WHEN LEN(@DateAsString)=17 AND @DateAsString LIKE '[0-9][0-9]-[0,1][0-9]-[0-3][0-9] [0-2][0-9]:[0-5][0-9]:[0-5][0-9]'											THEN 1 ELSE 0 END-- 20 = yy-mm-dd hh:mi:ss(24h) -- Stating 20 causes error!
			+CASE WHEN LEN(@DateAsString)>19 AND @DateAsString LIKE '[0-2][0-9][0-9][0-9]-[0,1][0-9]-[0-3][0-9] [0-2][0-9]:[0-5][0-9]:[0-5][0-9].0000000'						THEN 1 ELSE 0 END--120 = yyyy-mm-dd hh:mi:ss(24h)
			+CASE WHEN @DateAsString IS NOT NULL AND ISDATE(@DateAsString)=0 AND @DateAsString < 'a' AND @DateAsString NOT LIKE '%-%' AND @DateAsString LIKE '%.%'				THEN 1 ELSE 0 END
			+CASE WHEN @DateAsString  LIKE '[0-9][0-9][0-9][0-9][0-9]'																											THEN 1 ELSE 0 END
			 as tinyint) = 0
			THEN CAST(0 as float)
		ELSE CAST(1 as float)--DateConversionAccuracy
 / CAST(-- number of possibilities ...
 CASE WHEN LEN(@DateAsString)=26 AND @DateAsString LIKE '[a-Z][a-Z][a-Z] [0-3][0-9] [1,2][0-9][0-9][0-9] [0,1][0-9]:[0-5][0-9]:[0-5][0-9]:[0-9][0-9][0-9][A,P][M]'	THEN 1 ELSE 0 END--109 = mon dd yyyy hh:mi:ss:mmmAM (or PM)
+CASE WHEN LEN(@DateAsString)=24 AND @DateAsString LIKE '[a-Z][a-Z][a-Z] [0-3][0-9] [0-9][0-9] [0,1][0-9]:[0-5][0-9]:[0-5][0-9]:[0-9][0-9][0-9][A,P][M]'			THEN 1 ELSE 0 END--  9 = mon dd yy hh:mi:ss:mmmAM (or PM)
+CASE WHEN LEN(@DateAsString)=24 AND @DateAsString LIKE '[1,2][0-9][0-9][0-9]-[0,1][0-9]-[0-3][0-9]T[0-2][0-9]:[0-5][0-9]:[0-5][0-9].[0-9][0-9][0-9]Z'				THEN 1 ELSE 0 END--127 = yyyy-mm-ddThh:mi:ss.mmmZ
+CASE WHEN LEN(@DateAsString)=24 AND @DateAsString LIKE '[0-3][0-9] [a-Z][a-Z][a-Z] [1,2][0-9][0-9][0-9] [0-9][0-9]:[0-9][0-9]:[0-9][0-9]:[0-9][0-9][0-9]'			THEN 1 ELSE 0 END--113 = dd mon yyyy hh:mi:ss:mmm(24h)
+CASE WHEN LEN(@DateAsString)=20 AND @DateAsString LIKE '[0-3][0-9] [a-Z][a-Z][a-Z] [1,2][0-9][0-9][0-9] [0-9][0-9]:[0-9][0-9]:[0-9][0-9]'							THEN 1 ELSE 0 END--113 = dd mon yyyy hh:mi:ss:mmm(24h)
+CASE WHEN LEN(@DateAsString)=22 AND @DateAsString LIKE '[0-3][0-9] [a-Z][a-Z][a-Z] [0-9][0-9] [0-9][0-9]:[0-9][0-9]:[0-9][0-9]:[0-9][0-9][0-9]'					THEN 1 ELSE 0 END-- 13 = dd mon yy hh:mi:ss:mmm(24h)
+CASE WHEN LEN(@DateAsString)=23 AND @DateAsString LIKE '[1,2][0-9][0-9][0-9]-[0,1][0-9]-[0-3][0-9] [0-2][0-9]:[0-5][0-9]:[0-5][0-9].[0-9][0-9][0-9]'				THEN 1 ELSE 0 END--121 = yyyy-mm-dd hh:mi:ss.mmm(24h)
+CASE WHEN LEN(@DateAsString)=23 AND @DateAsString LIKE '[1,2][0-9][0-9][0-9]-[0,1][0-9]-[0-3][0-9]T[0-2][0-9]:[0-5][0-9]:[0-5][0-9].[0-9][0-9][0-9]'				THEN 1 ELSE 0 END--126 = yyyy-mm-ddThh:mi:ss.mmm (no spaces)
+CASE WHEN LEN(@DateAsString)=19 AND @DateAsString LIKE '[0-2][0-9][0-9][0-9]-[0,1][0-9]-[0-3][0-9] [0-2][0-9]:[0-5][0-9]:[0-5][0-9]'								THEN 1 ELSE 0 END--120 = yyyy-mm-dd hh:mi:ss(24h)
+CASE WHEN LEN(@DateAsString)=19 AND @DateAsString LIKE '[A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y] [0-3][0-9] [0-2][0-9][0-9][0-9] [0,1][0-9]:[0-5][0-9][A,P]M' THEN 1 ELSE 0 END--100 = mon dd yyyy hh:miAM (or PM)
+CASE WHEN LEN(@DateAsString)=17 AND @DateAsString LIKE '[A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y] [0-3][0-9] [0-9][0-9] [0,1][0-9]:[0-5][0-9][A,P]M'		THEN 1 ELSE 0 END--  0 = mon dd yy hh:miAM (or PM)
+CASE WHEN LEN(@DateAsString)=12 AND @DateAsString LIKE '[0-2][0-9]:[0-5][0-9]:[0-5][0-9]:[0-9][0-9][0-9]'															THEN 1 ELSE 0 END--114 = hh:mi:ss:mmm(24h)
+CASE WHEN LEN(@DateAsString)=10 AND SUBSTRING(@DateAsString,4,2) < '13'
									AND @DateAsString LIKE '[0-3][0-9]/[0,1][0-9]/[0-2][0-9][0-9][0-9]'																THEN 1 ELSE 0 END--103 = dd/mm/yyyy
+CASE WHEN LEN(@DateAsString)=16 AND @DateAsString LIKE '[0-3][0-9]/[0,1][0-9]/[0-2][0-9][0-9][0-9] [0-2][0-9]:[0-5][0-9]'											THEN 1 ELSE 0 END --103+108
+CASE WHEN LEN(@DateAsString)= 8 AND SUBSTRING(@DateAsString,1,2) < '32' AND SUBSTRING(@DateAsString,4,2) < '13' 
								AND @DateAsString LIKE '[0-3][0-9]/[0,1][0-9]/[0-9][0-9]'																			THEN 1 ELSE 0 END--  3 = dd/mm/yy
+CASE WHEN LEN(@DateAsString)=10 AND SUBSTRING(@DateAsString,1,2) < '13'
									AND @DateAsString LIKE '[0,1][0-9]/[0-3][0-9]/[0-2][0-9][0-9][0-9]'																THEN 1 ELSE 0 END--101 = mm/dd/yyyy
+CASE WHEN LEN(@DateAsString)=20 AND SUBSTRING(@DateAsString,1,2) < '13'
									AND @DateAsString LIKE '[0,1][0-9]/[0-3][0-9]/[0-9][0-9] [ ,0,1][0-9]:[0-5][0-9]:[0-5][0-9] [A,P]M'								THEN 1 ELSE 0 END-- 22 = mm/dd/yy hh:mi:ss AM (or PM)
+CASE WHEN LEN(@DateAsString)=10 AND @DateAsString LIKE '[0-2][0-9][0-9][0-9]/[0,1][0-9]/[0-3][0-9]'																THEN 1 ELSE 0 END--111 = yyyy/mm/dd
+CASE WHEN LEN(@DateAsString)= 8 AND SUBSTRING(@DateAsString,4,2) < '13' AND SUBSTRING(@DateAsString,6,2) < '32' 
								AND @DateAsString LIKE '[0-9][0-9]/[0,1][0-9]/[0-3][0-9]'																			THEN 1 ELSE 0 END-- 11 = yy/mm/dd
+CASE WHEN LEN(@DateAsString)= 8 AND SUBSTRING(@DateAsString,1,2) < '13' AND SUBSTRING(@DateAsString,4,2) < '32' 
								AND @DateAsString LIKE '[0,1][0-9]/[0-3][0-9]/[0-9][0-9]'																			THEN 1 ELSE 0 END--  1 = mm/dd/yy
+CASE WHEN LEN(@DateAsString)=10 AND @DateAsString LIKE '[A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y] [0-3][0-9], [0-9][0-9]'								THEN 1 ELSE 0 END--  7 = Mon dd, yy
+CASE WHEN LEN(@DateAsString)=12 AND @DateAsString LIKE '[A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y] [0-3][0-9], [0-2][0-9][0-9][0-9]'						THEN 1 ELSE 0 END--107 = Mon dd, yyyy
+CASE WHEN LEN(@DateAsString)= 9 AND @DateAsString LIKE '[0-3][0-9][ ,-][A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y][ ,-][0-9][0-9]'							THEN 1 ELSE 0 END--  6 = dd mon yy
+CASE WHEN LEN(@DateAsString)=11 AND @DateAsString LIKE '[0-3][0-9][ ,-][A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y][ ,-][0-2][0-9][0-9][0-9]'				THEN 1 ELSE 0 END--106 = dd mon yyyy
+CASE WHEN LEN(@DateAsString)= 8 AND @DateAsString LIKE '[0-9][0-9].[0,1][0-9].[0-3][0-9]'																			THEN 1 ELSE 0 END--  2 = yy.mm.dd
+CASE WHEN LEN(@DateAsString)=10 AND @DateAsString LIKE '[0-2][0-9][0-9][0-9].[0,1][0-9].[0-3][0-9]'																THEN 1 ELSE 0 END--102 = yyyy.mm.dd
+CASE WHEN LEN(@DateAsString)= 8 AND @DateAsString LIKE '[0-3][0-9].[0,1][0-9].[0-9][0-9]'																			THEN 1 ELSE 0 END--  4 = dd.mm.yy
+CASE WHEN LEN(@DateAsString)=10 AND @DateAsString LIKE '[0-3][0-9].[0,1][0-9].[0-2][0-9][0-9][0-9]'																THEN 1 ELSE 0 END--104 = dd.mm.yyyy
+CASE WHEN LEN(@DateAsString)= 8 AND @DateAsString LIKE '[0-3][0-9]-[0,1][0-9]-[0-9][0-9]'																			THEN 1 ELSE 0 END--  5 = dd-mm-yy
+CASE WHEN LEN(@DateAsString)=10 AND @DateAsString LIKE '[0-3][0-9]-[0,1][0-9]-[0-2][0-9][0-9][0-9]'																THEN 1 ELSE 0 END--105 = dd-mm-yyyy
+CASE WHEN LEN(@DateAsString)= 8 AND @DateAsString LIKE '[0-2][0-9]:[0-5][0-9]:[0-5][0-9]'																			THEN 1 ELSE 0 END--108 = hh:mi:ss
+CASE WHEN LEN(@DateAsString)= 7 AND @DateAsString LIKE '[0-9]:[0-5][0-9]:[0-5][0-9]'																				THEN 1 ELSE 0 END--108 = hh:mi:ss - leading zero truncated
+CASE WHEN LEN(@DateAsString)= 8 AND @DateAsString LIKE '[1,2][0-9][0-9][0-9][0,1][0-9][0-3][0-9]'																	THEN 1 ELSE 0 END--112 = yyyymmdd
+CASE WHEN LEN(@DateAsString)= 8 AND @DateAsString LIKE '[0,1][0-9]-[0-3][0-9]-[0-9][0-9]'																			THEN 1 ELSE 0 END-- 10 = mm-dd-yy
+CASE WHEN LEN(@DateAsString)= 6 AND @DateAsString LIKE '[0-9][0-9][0,1][0-9][0-3][0-9]'																			THEN 1 ELSE 0 END-- 12 = yymmdd
+CASE WHEN LEN(@DateAsString)=26 AND @DateAsString LIKE '[0-3][0-9] [a-Z][a-Z][a-Z] [1,2][0-9][0-9][0-9] [0,1][0-9]:[0-5][0-9]:[0-5][0-9]:[0-9][0-9][0-9][A,P][M]'	THEN 1 ELSE 0 END--130 = dd mon yyyy hh:mi:ss:mmmAM -- Stating 130 causes error!
--+CASE WHEN LEN(@DateAsString)=23 AND @DateAsString LIKE '[0-3][0-9]/[0,1][0-9]/[0-9][0-9] [0,1][0-9]:[0-5][0-9]:[0-5][0-9]:[0-9][0-9][0-9][A,P]M'					THEN 1  ELSE 0 END--131 = dd/mm/yy hh:mi:ss:mmmAM -- Bombs out with error!
+CASE WHEN LEN(@DateAsString)=21 AND @DateAsString LIKE '[0-9][0-9]-[0,1][0-9]-[0-3][0-9] [0-2][0-9]:[0-5][0-9]:[0-5][0-9].[0-9][0-9][0-9]'							THEN 1 ELSE 0 END-- 21 = yy-mm-dd hh:mi:ss.mmm(24h) -- Stating 21 causes error!
+CASE WHEN LEN(@DateAsString)=17 AND @DateAsString LIKE '[0-9][0-9]-[0,1][0-9]-[0-3][0-9] [0-2][0-9]:[0-5][0-9]:[0-5][0-9]'											THEN 1 ELSE 0 END-- 20 = yy-mm-dd hh:mi:ss(24h) -- Stating 20 causes error!
+CASE WHEN LEN(@DateAsString)>19 AND @DateAsString LIKE '[0-2][0-9][0-9][0-9]-[0,1][0-9]-[0-3][0-9] [0-2][0-9]:[0-5][0-9]:[0-5][0-9].0000000'						THEN 1 ELSE 0 END--120 = yyyy-mm-dd hh:mi:ss(24h)
+CASE WHEN @DateAsString IS NOT NULL AND ISDATE(@DateAsString)=0 AND @DateAsString < 'a' AND @DateAsString NOT LIKE '%-%' AND @DateAsString LIKE '%.%'				THEN 1 ELSE 0 END
+CASE WHEN @DateAsString  LIKE '[0-9][0-9][0-9][0-9][0-9]'																											THEN 1 ELSE 0 END
 as float) END,5) as DateConversionAccuracy

,CAST(--NumberOfPossibleTargetCodes
 CASE WHEN LEN(@DateAsString)=26 AND @DateAsString LIKE '[a-Z][a-Z][a-Z] [0-3][0-9] [1,2][0-9][0-9][0-9] [0,1][0-9]:[0-5][0-9]:[0-5][0-9]:[0-9][0-9][0-9][A,P][M]'	THEN 1 ELSE 0 END--109 = mon dd yyyy hh:mi:ss:mmmAM (or PM)
+CASE WHEN LEN(@DateAsString)=24 AND @DateAsString LIKE '[a-Z][a-Z][a-Z] [0-3][0-9] [0-9][0-9] [0,1][0-9]:[0-5][0-9]:[0-5][0-9]:[0-9][0-9][0-9][A,P][M]'			THEN 1 ELSE 0 END--  9 = mon dd yy hh:mi:ss:mmmAM (or PM)
+CASE WHEN LEN(@DateAsString)=24 AND @DateAsString LIKE '[1,2][0-9][0-9][0-9]-[0,1][0-9]-[0-3][0-9]T[0-2][0-9]:[0-5][0-9]:[0-5][0-9].[0-9][0-9][0-9]Z'				THEN 1 ELSE 0 END--127 = yyyy-mm-ddThh:mi:ss.mmmZ
+CASE WHEN LEN(@DateAsString)=24 AND @DateAsString LIKE '[0-3][0-9] [a-Z][a-Z][a-Z] [1,2][0-9][0-9][0-9] [0-9][0-9]:[0-9][0-9]:[0-9][0-9]:[0-9][0-9][0-9]'			THEN 1 ELSE 0 END--113 = dd mon yyyy hh:mi:ss:mmm(24h)
+CASE WHEN LEN(@DateAsString)=20 AND @DateAsString LIKE '[0-3][0-9] [a-Z][a-Z][a-Z] [1,2][0-9][0-9][0-9] [0-9][0-9]:[0-9][0-9]:[0-9][0-9]'							THEN 1 ELSE 0 END--113 = dd mon yyyy hh:mi:ss:mmm(24h)
+CASE WHEN LEN(@DateAsString)=22 AND @DateAsString LIKE '[0-3][0-9] [a-Z][a-Z][a-Z] [0-9][0-9] [0-9][0-9]:[0-9][0-9]:[0-9][0-9]:[0-9][0-9][0-9]'					THEN 1 ELSE 0 END-- 13 = dd mon yy hh:mi:ss:mmm(24h)
+CASE WHEN LEN(@DateAsString)=23 AND @DateAsString LIKE '[1,2][0-9][0-9][0-9]-[0,1][0-9]-[0-3][0-9] [0-2][0-9]:[0-5][0-9]:[0-5][0-9].[0-9][0-9][0-9]'				THEN 1 ELSE 0 END--121 = yyyy-mm-dd hh:mi:ss.mmm(24h)
+CASE WHEN LEN(@DateAsString)=23 AND @DateAsString LIKE '[1,2][0-9][0-9][0-9]-[0,1][0-9]-[0-3][0-9]T[0-2][0-9]:[0-5][0-9]:[0-5][0-9].[0-9][0-9][0-9]'				THEN 1 ELSE 0 END--126 = yyyy-mm-ddThh:mi:ss.mmm (no spaces)
+CASE WHEN LEN(@DateAsString)=19 AND @DateAsString LIKE '[0-2][0-9][0-9][0-9]-[0,1][0-9]-[0-3][0-9] [0-2][0-9]:[0-5][0-9]:[0-5][0-9]'								THEN 1 ELSE 0 END--120 = yyyy-mm-dd hh:mi:ss(24h)
+CASE WHEN LEN(@DateAsString)=19 AND @DateAsString LIKE '[A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y] [0-3][0-9] [0-2][0-9][0-9][0-9] [0,1][0-9]:[0-5][0-9][A,P]M' THEN 1 ELSE 0 END--100 = mon dd yyyy hh:miAM (or PM)
+CASE WHEN LEN(@DateAsString)=17 AND @DateAsString LIKE '[A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y] [0-3][0-9] [0-9][0-9] [0,1][0-9]:[0-5][0-9][A,P]M'		THEN 1 ELSE 0 END--  0 = mon dd yy hh:miAM (or PM)
+CASE WHEN LEN(@DateAsString)=12 AND @DateAsString LIKE '[0-2][0-9]:[0-5][0-9]:[0-5][0-9]:[0-9][0-9][0-9]'															THEN 1 ELSE 0 END--114 = hh:mi:ss:mmm(24h)
+CASE WHEN LEN(@DateAsString)=10 AND SUBSTRING(@DateAsString,4,2) < '13'
									AND @DateAsString LIKE '[0-3][0-9]/[0,1][0-9]/[0-2][0-9][0-9][0-9]'																THEN 1 ELSE 0 END--103 = dd/mm/yyyy
+CASE WHEN LEN(@DateAsString)=16 AND @DateAsString LIKE '[0-3][0-9]/[0,1][0-9]/[0-2][0-9][0-9][0-9] [0-2][0-9]:[0-5][0-9]'											THEN 1 ELSE 0 END --103+108
+CASE WHEN LEN(@DateAsString)= 8 AND SUBSTRING(@DateAsString,1,2) < '32' AND SUBSTRING(@DateAsString,4,2) < '13' 
								AND @DateAsString LIKE '[0-3][0-9]/[0,1][0-9]/[0-9][0-9]'																			THEN 1 ELSE 0 END--  3 = dd/mm/yy
+CASE WHEN LEN(@DateAsString)=10 AND SUBSTRING(@DateAsString,1,2) < '13'
									AND @DateAsString LIKE '[0,1][0-9]/[0-3][0-9]/[0-2][0-9][0-9][0-9]'																THEN 1 ELSE 0 END--101 = mm/dd/yyyy
+CASE WHEN LEN(@DateAsString)=20 AND SUBSTRING(@DateAsString,1,2) < '13'
									AND @DateAsString LIKE '[0,1][0-9]/[0-3][0-9]/[0-9][0-9] [ ,0,1][0-9]:[0-5][0-9]:[0-5][0-9] [A,P]M'								THEN 1 ELSE 0 END-- 22 = mm/dd/yy hh:mi:ss AM (or PM)
+CASE WHEN LEN(@DateAsString)=10 AND @DateAsString LIKE '[0-2][0-9][0-9][0-9]/[0,1][0-9]/[0-3][0-9]'																THEN 1 ELSE 0 END--111 = yyyy/mm/dd
+CASE WHEN LEN(@DateAsString)= 8 AND SUBSTRING(@DateAsString,4,2) < '13' AND SUBSTRING(@DateAsString,6,2) < '32' 
								AND @DateAsString LIKE '[0-9][0-9]/[0,1][0-9]/[0-3][0-9]'																			THEN 1 ELSE 0 END-- 11 = yy/mm/dd
+CASE WHEN LEN(@DateAsString)= 8 AND SUBSTRING(@DateAsString,1,2) < '13' AND SUBSTRING(@DateAsString,4,2) < '32' 
								AND @DateAsString LIKE '[0,1][0-9]/[0-3][0-9]/[0-9][0-9]'																			THEN 1 ELSE 0 END--  1 = mm/dd/yy
+CASE WHEN LEN(@DateAsString)=10 AND @DateAsString LIKE '[A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y] [0-3][0-9], [0-9][0-9]'								THEN 1 ELSE 0 END--  7 = Mon dd, yy
+CASE WHEN LEN(@DateAsString)=12 AND @DateAsString LIKE '[A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y] [0-3][0-9], [0-2][0-9][0-9][0-9]'						THEN 1 ELSE 0 END--107 = Mon dd, yyyy
+CASE WHEN LEN(@DateAsString)= 9 AND @DateAsString LIKE '[0-3][0-9][ ,-][A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y][ ,-][0-9][0-9]'							THEN 1 ELSE 0 END--  6 = dd mon yy
+CASE WHEN LEN(@DateAsString)=11 AND @DateAsString LIKE '[0-3][0-9][ ,-][A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y][ ,-][0-2][0-9][0-9][0-9]'				THEN 1 ELSE 0 END--106 = dd mon yyyy
+CASE WHEN LEN(@DateAsString)= 8 AND @DateAsString LIKE '[0-9][0-9].[0,1][0-9].[0-3][0-9]'																			THEN 1 ELSE 0 END--  2 = yy.mm.dd
+CASE WHEN LEN(@DateAsString)=10 AND @DateAsString LIKE '[0-2][0-9][0-9][0-9].[0,1][0-9].[0-3][0-9]'																THEN 1 ELSE 0 END--102 = yyyy.mm.dd
+CASE WHEN LEN(@DateAsString)= 8 AND @DateAsString LIKE '[0-3][0-9].[0,1][0-9].[0-9][0-9]'																			THEN 1 ELSE 0 END--  4 = dd.mm.yy
+CASE WHEN LEN(@DateAsString)=10 AND @DateAsString LIKE '[0-3][0-9].[0,1][0-9].[0-2][0-9][0-9][0-9]'																THEN 1 ELSE 0 END--104 = dd.mm.yyyy
+CASE WHEN LEN(@DateAsString)= 8 AND @DateAsString LIKE '[0-3][0-9]-[0,1][0-9]-[0-9][0-9]'																			THEN 1 ELSE 0 END--  5 = dd-mm-yy
+CASE WHEN LEN(@DateAsString)=10 AND @DateAsString LIKE '[0-3][0-9]-[0,1][0-9]-[0-2][0-9][0-9][0-9]'																THEN 1 ELSE 0 END--105 = dd-mm-yyyy
+CASE WHEN LEN(@DateAsString)= 8 AND @DateAsString LIKE '[0-2][0-9]:[0-5][0-9]:[0-5][0-9]'																			THEN 1 ELSE 0 END--108 = hh:mi:ss
+CASE WHEN LEN(@DateAsString)= 7 AND @DateAsString LIKE '[0-9]:[0-5][0-9]:[0-5][0-9]'																				THEN 1 ELSE 0 END--108 = hh:mi:ss - leading zero truncated
+CASE WHEN LEN(@DateAsString)= 8 AND @DateAsString LIKE '[1,2][0-9][0-9][0-9][0,1][0-9][0-3][0-9]'																	THEN 1 ELSE 0 END--112 = yyyymmdd
+CASE WHEN LEN(@DateAsString)= 8 AND @DateAsString LIKE '[0,1][0-9]-[0-3][0-9]-[0-9][0-9]'																			THEN 1 ELSE 0 END-- 10 = mm-dd-yy
+CASE WHEN LEN(@DateAsString)= 6 AND @DateAsString LIKE '[0-9][0-9][0,1][0-9][0-3][0-9]'																			THEN 1 ELSE 0 END-- 12 = yymmdd
+CASE WHEN LEN(@DateAsString)=26 AND @DateAsString LIKE '[0-3][0-9] [a-Z][a-Z][a-Z] [1,2][0-9][0-9][0-9] [0,1][0-9]:[0-5][0-9]:[0-5][0-9]:[0-9][0-9][0-9][A,P][M]'	THEN 1 ELSE 0 END--130 = dd mon yyyy hh:mi:ss:mmmAM -- Stating 130 causes error!
--+CASE WHEN LEN(@DateAsString)=23 AND @DateAsString LIKE '[0-3][0-9]/[0,1][0-9]/[0-9][0-9] [0,1][0-9]:[0-5][0-9]:[0-5][0-9]:[0-9][0-9][0-9][A,P]M'					THEN 1  ELSE 0 END--131 = dd/mm/yy hh:mi:ss:mmmAM -- Bombs out with error!
+CASE WHEN LEN(@DateAsString)=21 AND @DateAsString LIKE '[0-9][0-9]-[0,1][0-9]-[0-3][0-9] [0-2][0-9]:[0-5][0-9]:[0-5][0-9].[0-9][0-9][0-9]'							THEN 1 ELSE 0 END-- 21 = yy-mm-dd hh:mi:ss.mmm(24h) -- Stating 21 causes error!
+CASE WHEN LEN(@DateAsString)=17 AND @DateAsString LIKE '[0-9][0-9]-[0,1][0-9]-[0-3][0-9] [0-2][0-9]:[0-5][0-9]:[0-5][0-9]'											THEN 1 ELSE 0 END-- 20 = yy-mm-dd hh:mi:ss(24h) -- Stating 20 causes error!
+CASE WHEN LEN(@DateAsString)>19 AND @DateAsString LIKE '[0-2][0-9][0-9][0-9]-[0,1][0-9]-[0-3][0-9] [0-2][0-9]:[0-5][0-9]:[0-5][0-9].0000000'						THEN 1 ELSE 0 END--120 = yyyy-mm-dd hh:mi:ss(24h)
+CASE WHEN @DateAsString IS NOT NULL AND ISDATE(@DateAsString)=0 AND @DateAsString < 'a' AND @DateAsString NOT LIKE '%-%' AND @DateAsString LIKE '%.%'				THEN 1 ELSE 0 END
+CASE WHEN @DateAsString  LIKE '[0-9][0-9][0-9][0-9][0-9]'																											THEN 1 ELSE 0 END
 as tinyint) as NumberOfPossibleTargetDateCodes

,CAST(LTRIM(RTRIM(REPLACE(-- possible formats ...
 CASE WHEN @DateAsString LIKE '[a-Z][a-Z][a-Z] [0-3][0-9] [1,2][0-9][0-9][0-9] [0,1][0-9]:[0-5][0-9]:[0-5][0-9]:[0-9][0-9][0-9][A,P]M%' THEN '109' ELSE '' END+' '--109 = mon dd yyyy hh:mi:ss:mmmAM (or PM)
+CASE WHEN @DateAsString LIKE '[a-Z][a-Z][a-Z] [0-3][0-9] [0-9][0-9] [0,1][0-9]:[0-5][0-9]:[0-5][0-9]:[0-9][0-9][0-9][A,P]M%'			THEN   '9' ELSE '' END+' '--  9 = mon dd yy hh:mi:ss:mmmAM (or PM)
+CASE WHEN @DateAsString LIKE '[1,2][0-9][0-9][0-9]-[0,1][0-9]-[0-3][0-9]T[0-2][0-9]:[0-5][0-9]:[0-5][0-9].[0-9][0-9][0-9]Z%'			THEN '127' ELSE '' END+' '--127 = yyyy-mm-ddThh:mi:ss.mmmZ
+CASE WHEN @DateAsString LIKE '[0-3][0-9] [a-Z][a-Z][a-Z] [1,2][0-9][0-9][0-9] [0-9][0-9]:[0-9][0-9]:[0-9][0-9]%'						THEN '113' ELSE '' END+' '--113 = dd mon yyyy hh:mi:ss:mmm(24h)
--+CASE WHEN @DateAsString LIKE '[0-3][0-9] [a-Z][a-Z][a-Z] [1,2][0-9][0-9][0-9] [0-9][0-9]:[0-9][0-9]:[0-9][0-9]:[0-9][0-9][0-9]%'		THEN '113' ELSE '' END+' '--113 = dd mon yyyy hh:mi:ss:mmm(24h)
+CASE WHEN @DateAsString LIKE '[0-3][0-9] [a-Z][a-Z][a-Z] [0-9][0-9] [0-9][0-9]:[0-9][0-9]:[0-9][0-9]:[0-9][0-9][0-9]%'					THEN  '13' ELSE '' END+' '-- 13 = dd mon yy hh:mi:ss:mmm(24h)
+CASE WHEN @DateAsString LIKE '[1,2][0-9][0-9][0-9]-[0,1][0-9]-[0-3][0-9] [0-2][0-9]:[0-5][0-9]:[0-5][0-9].[0-9][0-9][0-9]%'			THEN '121' ELSE '' END+' '--121 = yyyy-mm-dd hh:mi:ss.mmm(24h)
+CASE WHEN @DateAsString LIKE '[1,2][0-9][0-9][0-9]-[0,1][0-9]-[0-3][0-9]T[0-2][0-9]:[0-5][0-9]:[0-5][0-9].[0-9][0-9][0-9]%'			THEN '126' ELSE '' END+' '--126 = yyyy-mm-ddThh:mi:ss.mmm (no spaces)
+CASE WHEN @DateAsString LIKE '[0-2][0-9][0-9][0-9]-[0,1][0-9]-[0-3][0-9] [0-2][0-9]:[0-5][0-9]:[0-5][0-9]%'							THEN '120' ELSE '' END+' '--120 = yyyy-mm-dd hh:mi:ss(24h)
+CASE WHEN @DateAsString LIKE '[A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y] [0-3][0-9] [0-2][0-9][0-9][0-9] [0,1][0-9]:[0-5][0-9][A,P]M%' THEN '100' ELSE '' END+' '--100 = mon dd yyyy hh:miAM (or PM)
+CASE WHEN @DateAsString LIKE '[A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y] [0-3][0-9] [0-9][0-9] [0,1][0-9]:[0-5][0-9][A,P]M%'	THEN   '0' ELSE '' END+' '--  0 = mon dd yy hh:miAM (or PM)
+CASE WHEN @DateAsString LIKE '[0-2][0-9]:[0-5][0-9]:[0-5][0-9]:[0-9][0-9][0-9]%'														THEN '114' ELSE '' END+' '--114 = hh:mi:ss:mmm(24h)
+CASE WHEN SUBSTRING(@DateAsString,4,2) < '13'
		AND @DateAsString LIKE '[0-3][0-9]/[0,1][0-9]/[0-2][0-9][0-9][0-9]%'															THEN '103' ELSE '' END+' '--103 = dd/mm/yyyy
+CASE WHEN @DateAsString LIKE '[0-3][0-9]/[0,1][0-9]/[0-2][0-9][0-9][0-9] [0-2][0-9]:[0-5][0-9]%'										THEN '211' ELSE '' END+' '--103+108
+CASE WHEN SUBSTRING(@DateAsString,1,2) < '32' AND SUBSTRING(@DateAsString,4,2) < '13' 
		AND @DateAsString LIKE '[0-3][0-9]/[0,1][0-9]/[0-9][0-9]%'																		THEN   '3' ELSE '' END+' '--  3 = dd/mm/yy
+CASE WHEN SUBSTRING(@DateAsString,1,2) < '13'
		AND @DateAsString LIKE '[0,1][0-9]/[0-3][0-9]/[0-2][0-9][0-9][0-9]%'															THEN '101' ELSE '' END+' '--101 = mm/dd/yyyy
+CASE WHEN SUBSTRING(@DateAsString,1,2) < '13'
		AND @DateAsString LIKE '[0,1][0-9]/[0-3][0-9]/[0-9][0-9] [ ,0,1][0-9]:[0-5][0-9]:[0-5][0-9] [A,P]M%'							THEN  '22' ELSE '' END+' '-- 22 = mm/dd/yy hh:mi:ss AM (or PM)
+CASE WHEN @DateAsString LIKE '[0-2][0-9][0-9][0-9]/[0,1][0-9]/[0-3][0-9]%'																THEN '111' ELSE '' END+' '--111 = yyyy/mm/dd
+CASE WHEN SUBSTRING(@DateAsString,4,2) < '13' AND SUBSTRING(@DateAsString,6,2) < '32' 
		AND @DateAsString LIKE '[0-9][0-9]/[0,1][0-9]/[0-3][0-9]%'																		THEN  '11' ELSE '' END+' '-- 11 = yy/mm/dd
+CASE WHEN SUBSTRING(@DateAsString,1,2) < '13' AND SUBSTRING(@DateAsString,4,2) < '32' 
		AND @DateAsString LIKE '[0,1][0-9]/[0-3][0-9]/[0-9][0-9]%'																		THEN   '1' ELSE '' END+' '--  1 = mm/dd/yy
+CASE WHEN @DateAsString LIKE '[A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y] [0-3][0-9], [0-9][0-9]%'								THEN   '7' ELSE '' END+' '--  7 = Mon dd, yy
+CASE WHEN @DateAsString LIKE '[A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y] [0-3][0-9], [0-2][0-9][0-9][0-9]%'					THEN '107' ELSE '' END+' '--107 = Mon dd, yyyy
+CASE WHEN @DateAsString LIKE '[0-3][0-9][ ,-][A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y][ ,-][0-9][0-9]%'						THEN   '6' ELSE '' END+' '--  6 = dd mon yy
+CASE WHEN @DateAsString LIKE '[0-3][0-9][ ,-][A,D,F,J,M,N,O,S][a,c,e,o,p,u][b,c,g,l,n,p,r,t,v,y][ ,-][0-2][0-9][0-9][0-9]%'			THEN '106' ELSE '' END+' '--106 = dd mon yyyy
+CASE WHEN @DateAsString LIKE '[0-9][0-9].[0,1][0-9].[0-3][0-9]%'																		THEN   '2' ELSE '' END+' '--  2 = yy.mm.dd
+CASE WHEN @DateAsString LIKE '[0-2][0-9][0-9][0-9].[0,1][0-9].[0-3][0-9]%'																THEN '102' ELSE '' END+' '--102 = yyyy.mm.dd
+CASE WHEN @DateAsString LIKE '[0-3][0-9].[0,1][0-9].[0-9][0-9]%'																		THEN   '4' ELSE '' END+' '--  4 = dd.mm.yy
+CASE WHEN @DateAsString LIKE '[0-3][0-9].[0,1][0-9].[0-2][0-9][0-9][0-9]%'																THEN '104' ELSE '' END+' '--104 = dd.mm.yyyy
+CASE WHEN @DateAsString LIKE '[0-3][0-9]-[0,1][0-9]-[0-9][0-9]%'																		THEN   '5' ELSE '' END+' '--  5 = dd-mm-yy
+CASE WHEN @DateAsString LIKE '[0-3][0-9]-[0,1][0-9]-[0-2][0-9][0-9][0-9]%'																THEN '105' ELSE '' END+' '--105 = dd-mm-yyyy
+CASE WHEN @DateAsString LIKE '[0-2][0-9]:[0-5][0-9]:[0-5][0-9]%'																		THEN '108' ELSE '' END+' '--108 = hh:mi:ss
+CASE WHEN @DateAsString LIKE '[0-9]:[0-5][0-9]:[0-5][0-9]'																				THEN '-108'ELSE '' END+' '--108 = hh:mi:ss - leading zero truncated
+CASE WHEN @DateAsString LIKE '[1,2][0-9][0-9][0-9][0,1][0-9][0-3][0-9]%'																THEN '112' ELSE '' END+' '--112 = yyyymmdd
+CASE WHEN @DateAsString LIKE '[0,1][0-9]-[0-3][0-9]-[0-9][0-9]%'																		THEN  '10' ELSE '' END+' '-- 10 = mm-dd-yy
+CASE WHEN @DateAsString LIKE '[0-9][0-9][0,1][0-9][0-3][0-9]%'																			THEN  '12' ELSE '' END+' '-- 12 = yymmdd
+CASE WHEN @DateAsString LIKE '[0-3][0-9] [a-Z][a-Z][a-Z] [1,2][0-9][0-9][0-9] [0,1][0-9]:[0-5][0-9]:[0-5][0-9]:[0-9][0-9][0-9][A,P]M%'	THEN '130' ELSE '' END+' '--130 = dd mon yyyy hh:mi:ss:mmmAM -- Stating 130 causes error!
+CASE WHEN @DateAsString LIKE '[0-3][0-9]/[0,1][0-9]/[0-9][0-9] [0,1][0-9]:[0-5][0-9]:[0-5][0-9]:[0-9][0-9][0-9][A,P]M%'				THEN '131' ELSE '' END+' '--131 = dd/mm/yy hh:mi:ss:mmmAM -- Bombs out with error!
+CASE WHEN @DateAsString LIKE '[0-9][0-9]-[0,1][0-9]-[0-3][0-9] [0-2][0-9]:[0-5][0-9]:[0-5][0-9].[0-9][0-9][0-9]%'						THEN  '21' ELSE '' END+' '-- 21 = yy-mm-dd hh:mi:ss.mmm(24h) -- Stating 21 causes error!
+CASE WHEN @DateAsString LIKE '[0-9][0-9]-[0,1][0-9]-[0-3][0-9] [0-2][0-9]:[0-5][0-9]:[0-5][0-9]%'										THEN  '20' ELSE '' END+' '-- 20 = yy-mm-dd hh:mi:ss(24h) -- Stating 20 causes error!
+CASE WHEN @DateAsString LIKE '[0-9]/[0-9]/[1,2][0-9][0-9][0-9] [0-9]:[0-5][0-9]'			  THEN 'Excel hell - zeros stripped from American DMH' ELSE '' END
+CASE WHEN @DateAsString LIKE '[0-9]/[0-9]/[1,2][0-9][0-9][0-9] [0-2][0-9]:[0-5][0-9]'		  THEN 'Excel hell - zeros stripped from American DM'  ELSE '' END
+CASE WHEN @DateAsString LIKE '[0-9]/[0-3][0-9]/[1,2][0-9][0-9][0-9] [0-9]:[0-5][0-9]'		  THEN 'Excel hell - zeros stripped from American DH'  ELSE '' END
+CASE WHEN @DateAsString LIKE '[0-9]/[0-3][0-9]/[1,2][0-9][0-9][0-9] [0-5][0-9]:[0-5][0-9]'	  THEN 'Excel hell - zeros stripped from American D'   ELSE '' END
+CASE WHEN @DateAsString IS NOT NULL AND ISDATE(@DateAsString)=0 AND @DateAsString < 'a' AND @DateAsString NOT LIKE '%-%' AND @DateAsString LIKE '%.%' THEN 'Excel hell - date as days (float) past 31-Dec-1899' ELSE '' END
+CASE WHEN @DateAsString  LIKE '[0-9][0-9][0-9][0-9][0-9]'																				THEN 'Excel hell - date as days (float) past 31-Dec-1899' ELSE '' END
,'  ',' '))) as varchar(50)) as PossibleFormats
)
GO