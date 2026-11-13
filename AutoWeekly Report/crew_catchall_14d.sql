-- ============================================================
-- Crew Packout Catchall Count — Last 14 Shift Days
-- Run against the plant's ODS server, database ODS:
--   PGT 1.1 / 1.4: PGT1MESSQLODS
--   PGT2:          PGT2MESSQLODS
--
-- A Catchall is a module put into ContainerId 'CATCHALL' at one of the
-- line's packout docks (SourceLocation like PGT14B-PACKOUT_DOCK4 or
-- PGT11A-PACKOUT_ROBOT_DOCK9), read from ODS.mfg.ProcessHistoryPdrContainerContentEvent.
-- The same rule works on every line, so there's no per-line branch.
-- It replaces the older rules (PACKOUT_INFEED matched to a CATCHALL
-- container event in ModuleAssembly, and the PGT 1.1 bin assignment),
-- which this table makes unnecessary.
--
-- A module sent to the Catchall more than once in the window counts
-- once, at its first Catchall event.
--
-- Window: the 14 shift days before today's shift start, shift days
-- starting at @ShiftStartHour, same as crew_frame_fg_catchall_14d.sql.
--
-- Speed: ods.sas.getcrew() is a scalar function, so it's called once
-- per hour (#CrewHour, ~340 calls) instead of once per module. Shifts
-- change on the hour, so each module gets the same crew as calling it
-- per module. OPTION (RECOMPILE) lets the optimizer see the real dates.
--
-- Edit @Plant before running: 'PGT11%', 'PGT14%' or 'PGT2%'.
--
-- Returns one row per crew:
--   Crew
--   Catchall   distinct modules sent to the Catchall
-- ============================================================

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
