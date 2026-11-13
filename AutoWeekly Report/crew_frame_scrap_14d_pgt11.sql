-- ============================================================
-- Crew Frame Scrap % — Last 14 Shift Days (PGT 1.1 ONLY)
-- Run against PGT1MESSQLODS, database ODS
--
-- Same source and logic as power_bi/ytd_frame_scrap_pct_pgt1.sql
-- (ProcessHistoryFrameAssyGlueStatus, the Adhesive/Glue reject rate
-- from "Frame Health Script v2.1.jsl"), rolled up to one row per crew.
-- PGT 1.4 and PGT2 don't log here: use crew_frame_scrap_14d_pgt14.sql /
-- crew_frame_scrap_14d_pgt2.sql for those.
--
-- Window: the 14 shift days before today's shift start, shift days
-- starting at @ShiftStartHour, same as crew_frame_fg_catchall_14d.sql.
--
-- Dedup: drops a glue event <2 seconds after the previous one on the
-- same FrameTool/EquipmentSide with TriggerCount exactly +1 (the JSL's
-- Repeat? column). Reads one minute before the window so the first
-- event in the window still has its previous event to compare against.
--
-- Returns one row per crew:
--   Crew
--   GlueTotal    glue-check events
--   GlueReject   events with GlueStatus = 'REJECT'
--   ScrapRate    GlueReject / GlueTotal
--   Target       0.8%
--   ActionLimit  1.0%
-- ============================================================

-- Day after the last day reported. Today by default; the report sets it
-- to the Monday after last week.
DECLARE @ReportEnd DATE = CAST(GETDATE() AS DATE);
DECLARE @ShiftStartHour      INT          = 6;
DECLARE @WindowDays          INT          = 14;
DECLARE @ScrapTargetPct      DECIMAL(6,4) = 0.0080;
DECLARE @ScrapActionLimitPct DECIMAL(6,4) = 0.0100;

DECLARE @End   DATETIME2 = DATEADD(HOUR, @ShiftStartHour, CAST(@ReportEnd AS DATETIME2));
DECLARE @Start DATETIME2 = DATEADD(DAY, -@WindowDays, @End);

WITH GlueRaw AS (
    SELECT
        EquipmentName,
        EquipmentSide,
        ReadTime,
        TriggerCount,
        GlueStatus,
        PrevReadTime     = LAG(ReadTime)     OVER (PARTITION BY EquipmentName, EquipmentSide ORDER BY ReadTime),
        PrevTriggerCount = LAG(TriggerCount) OVER (PARTITION BY EquipmentName, EquipmentSide ORDER BY ReadTime)
    FROM ODS.mfg.ProcessHistoryFrameAssyGlueStatus
    WHERE ReadTime >= DATEADD(MINUTE, -1, @Start)
      AND ReadTime <  @End
      AND SUBSTRING(EquipmentName, 5, 1) = '1'
),
GlueDedup AS (
    SELECT *
    FROM GlueRaw
    WHERE ReadTime >= @Start
      AND NOT (
        PrevReadTime IS NOT NULL
        AND DATEDIFF(SECOND, PrevReadTime, ReadTime) < 2
        AND TriggerCount - PrevTriggerCount = 1
      )
)
SELECT
    Crew       = [ods].[sas].[getcrew](ods.mfg.fn_plant(), ReadTime),
    GlueTotal  = COUNT(*),
    GlueReject = SUM(CASE WHEN GlueStatus = 'REJECT' THEN 1 ELSE 0 END),
    ScrapRate  = CONVERT(DECIMAL(10,5),
                 SUM(CASE WHEN GlueStatus = 'REJECT' THEN 1 ELSE 0 END) * 1.0 / NULLIF(COUNT(*), 0)),
    Target      = @ScrapTargetPct,
    ActionLimit = @ScrapActionLimitPct
FROM GlueDedup
GROUP BY
    [ods].[sas].[getcrew](ods.mfg.fn_plant(), ReadTime)
ORDER BY
    Crew;
