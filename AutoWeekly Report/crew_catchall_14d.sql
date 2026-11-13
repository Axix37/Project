SET NOCOUNT ON;

-- Day after the last day reported. Today by default; the report sets it
-- to the Monday after last week.
DECLARE @ReportEnd DATE = CAST(GETDATE() AS DATE);
DECLARE @Plant          VARCHAR(10) = 'PGT11%';
DECLARE @ShiftStartHour INT         = 6;
DECLARE @WindowDays     INT         = 14;

DECLARE @End   DATETIME2 = DATEADD(HOUR, @ShiftStartHour, CAST(@ReportEnd AS DATETIME2));
DECLARE @Start DATETIME2 = DATEADD(DAY, -@WindowDays, @End);

IF OBJECT_ID('tempdb..#Catchall') IS NOT NULL DROP TABLE #Catchall;
IF OBJECT_ID('tempdb..#CrewHour') IS NOT NULL DROP TABLE #CrewHour;

-- First Catchall event per module in the window.
SELECT
    ID        = cce.SubId,
    TimeStamp = MIN(cce.TimeStamp)
INTO #Catchall
FROM ODS.mfg.ProcessHistoryPdrContainerContentEvent cce
WHERE cce.TimeStamp >= @Start
  AND cce.TimeStamp <  @End
  AND cce.ContainerId = 'CATCHALL'
  AND cce.SourceLocation LIKE @Plant
  AND cce.SourceLocation LIKE '%PACKOUT%DOCK%'
  AND cce.SubId <> ''
GROUP BY cce.SubId
OPTION (RECOMPILE);

-- Crew on shift for each hour of the window.
WITH Hours AS (
    SELECT HourStart = @Start
    UNION ALL
    SELECT DATEADD(HOUR, 1, HourStart) FROM Hours WHERE DATEADD(HOUR, 1, HourStart) < @End
)
SELECT
    HourStart,
    Crew = [ods].[sas].[getcrew](ods.mfg.fn_plant(), HourStart)
INTO #CrewHour
FROM Hours
OPTION (MAXRECURSION 1000);

CREATE CLUSTERED INDEX IX_CrewHour ON #CrewHour (HourStart);

SELECT
    Crew     = H.Crew,
    Catchall = COUNT(DISTINCT P.ID)
FROM #Catchall P
INNER JOIN #CrewHour H
    ON H.HourStart = DATEADD(HOUR, DATEDIFF(HOUR, '19000101', P.TimeStamp), CAST('19000101' AS DATETIME2))
GROUP BY
    H.Crew
ORDER BY
    H.Crew;

DROP TABLE #Catchall;
DROP TABLE #CrewHour;
