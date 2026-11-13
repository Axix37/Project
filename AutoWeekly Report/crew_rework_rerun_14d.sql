

SET NOCOUNT ON;

-- Day after the last day reported. Today by default; the report sets it
-- to the Monday after last week.
DECLARE @ReportEnd DATE = CAST(GETDATE() AS DATE);
DECLARE @Plant          VARCHAR(10) = 'PGT11%';
DECLARE @ShiftStartHour INT         = 6;
DECLARE @WindowDays     INT         = 14;

DECLARE @End   DATETIME = DATEADD(HOUR, @ShiftStartHour, CAST(@ReportEnd AS DATETIME));
DECLARE @Start DATETIME = DATEADD(DAY, -@WindowDays, @End);
DECLARE @UtcStart DATETIME = DATEADD(HOUR, -12, @Start);
DECLARE @UtcEnd   DATETIME = DATEADD(HOUR,  12, @End);

IF OBJECT_ID('tempdb..#Tested') IS NOT NULL DROP TABLE #Tested;
IF OBJECT_ID('tempdb..#Rerun')  IS NOT NULL DROP TABLE #Rerun;

-- Modules tested per unit and hour.
SELECT
    Site          = SUBSTRING(TP.SitePlant, 1, 4),
    SitePlantUnit = SUBSTRING(TP.SourceLocation, 1, 5),
    TP.ReadTimeHour,
    Tested        = SUM(TP.PartProduced)
INTO #Tested
FROM MesReporting.IntRi.RunTimeHipotTPV2Global AS TP
WHERE TP.ReadTimeHourUtc >= @UtcStart
  AND TP.ReadTimeHourUtc <  @UtcEnd
  AND TP.ReadTimeHour    >= @Start
  AND TP.ReadTimeHour    <  @End
  AND TP.SourceLocation LIKE @Plant
GROUP BY
    SUBSTRING(TP.SitePlant, 1, 4),
    SUBSTRING(TP.SourceLocation, 1, 5),
    TP.ReadTimeHour
OPTION (RECOMPILE);

-- Reruns per unit and hour.
SELECT
    SitePlantUnit = SUBSTRING(RR.SourceLocation, 1, 5),
    ReadTimeHour  = DATEADD(HOUR, DATEPART(HOUR, RR.[TimeStamp]), CAST(CAST(RR.[TimeStamp] AS DATE) AS DATETIME)),
    Rerun         = COUNT(*)
INTO #Rerun
FROM MesReporting.IntRi.RunTimeHipotRerunV2Global AS RR
WHERE RR.ReadTimeHourUtc >= @UtcStart
  AND RR.ReadTimeHourUtc <  @UtcEnd
  AND RR.[TimeStamp]     >= @Start
  AND RR.[TimeStamp]     <  @End
  AND RR.SourceLocation LIKE @Plant
  AND RR.flag_calibration IS NULL
  AND RR.flag_excluding   IS NULL
GROUP BY
    SUBSTRING(RR.SourceLocation, 1, 5),
    DATEADD(HOUR, DATEPART(HOUR, RR.[TimeStamp]), CAST(CAST(RR.[TimeStamp] AS DATE) AS DATETIME))
OPTION (RECOMPILE);

-- Same LEFT JOIN as the Power BI query: reruns count only in hours the
-- unit tested modules.
WITH Hourly AS (
    SELECT
        Crew   = ODS.sas.GetCrew(T.Site, T.ReadTimeHour),
        T.Tested,
        Rerun  = COALESCE(R.Rerun, 0)
    FROM #Tested T
    LEFT JOIN #Rerun R
        ON  R.SitePlantUnit = T.SitePlantUnit
        AND R.ReadTimeHour  = T.ReadTimeHour
)
SELECT
    Crew,
    Tested      = SUM(Tested),
    Rerun       = SUM(Rerun),
    ReworkRerun = CONVERT(DECIMAL(10,5), SUM(Rerun) * 1.0 / NULLIF(SUM(Tested), 0))
FROM Hourly
GROUP BY Crew
ORDER BY Crew;
