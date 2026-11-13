
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
