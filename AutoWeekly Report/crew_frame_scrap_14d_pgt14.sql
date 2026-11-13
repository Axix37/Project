-- ============================================================
-- Crew Frame Scrap % — Last 14 Shift Days (PGT 1.4 ONLY)
-- Run against PGT1MESSQLODS, database ODS
--
-- Same source as power_bi/ytd_frame_scrap_pct_pgt14.sql
-- (ProcessHistoryPdrPartProducedEvent at FRAME_ASSEMBLY, decoded from
-- ExitProductStatus), rolled up to one row per crew.
-- PGT 1.1 logs frame scrap in a different table (crew_frame_scrap_14d_pgt11.sql).
--
-- Window: the 14 shift days before today's shift start, shift days
-- starting at @ShiftStartHour, same as crew_frame_fg_catchall_14d.sql.
--
-- Classification: frame ADHESIVE reject, as in "PGT14 FIN3 Framing
-- Datav0.1.jsl" (FlagYield = FAIL_ADHESIVE, REJECT ADHESIVE column):
--   FrameReject: FAIL_VISIONX ('500'), FAIL_RATIO_VISIONX ('510'),
--                FAIL_VISIONY ('600'), FAIL_RATIO_VISIONY ('610'),
--                FAIL_SILICONE_EXPIRE ('800')
--   FrameTotal:  GOOD_SCREW ('100') + GOOD_REWORK ('200') + FrameReject
--   Not counted: screw, DVM, tape, initialize, Sika and other FAIL_*
--                codes, MISS_FRAME_DATA* and the peel test sample ('850')
--
-- Returns one row per crew:
--   Crew
--   FrameTotal
--   FrameReject
--   ScrapRate    FrameReject / FrameTotal (adhesive reject %)
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

WITH FrameEvents AS (
    SELECT
        TimeStamp,
        FrameStatus = CASE ExitProductStatus
            WHEN '000'  THEN 'MISS_FRAME_DATA'
            WHEN '1000' THEN 'MISS_FRAME_DATA_BUFFER'
            WHEN '1001' THEN 'MISS_FRAME_DATA_FLIPPER'
            WHEN '1002' THEN 'MISS_FRAME_DATA_DVM'
            WHEN '1003' THEN 'MISS_FRAME_DATA_TAPE_X1'
            WHEN '1004' THEN 'MISS_FRAME_DATA_TAPE_X2'
            WHEN '1005' THEN 'MISS_FRAME_DATA_TAPE_Y'
            WHEN '1006' THEN 'MISS_FRAME_DATA_TAPE_Y_BUFFER'
            WHEN '1007' THEN 'MISS_FRAME_DATA_SILICONE_X'
            WHEN '1008' THEN 'MISS_FRAME_DATA_SILICONE_Y'
            WHEN '1009' THEN 'MISS_FRAME_DATA_BUFFER_POST_SILICONE'
            WHEN '1020' THEN 'MISS_FRAME_DATA_PAIRING'
            WHEN '100'  THEN 'GOOD_SCREW'
            WHEN '150'  THEN 'FAIL_SCREW'
            WHEN '200'  THEN 'GOOD_REWORK'
            WHEN '300'  THEN 'FAIL_INITIALIZE_BUFFER'
            WHEN '301'  THEN 'FAIL_INITIALIZE_FLIPPER'
            WHEN '302'  THEN 'FAIL_INITIALIZE_DVM'
            WHEN '303'  THEN 'FAIL_INITIALIZE_TAPE_X1'
            WHEN '304'  THEN 'FAIL_INITIALIZE_TAPE_X2'
            WHEN '305'  THEN 'FAIL_INITIALIZE_TAPE_Y'
            WHEN '306'  THEN 'FAIL_INITIALIZE_TAPE_Y_BUFFER'
            WHEN '307'  THEN 'FAIL_INITIALIZE_SILICONE_X'
            WHEN '308'  THEN 'FAIL_INITIALIZE_SILICONE_Y'
            WHEN '309'  THEN 'FAIL_INITIALIZE_BUFFER_POST_SILICONE'
            WHEN '321'  THEN 'FAIL_INITIALIZE_PAIRING'
            WHEN '350'  THEN 'FAIL_SIKA'
            WHEN '400'  THEN 'FAIL_DVM'
            WHEN '425'  THEN 'FAIL_TAPE_X1'
            WHEN '450'  THEN 'FAIL_TAPE_X2'
            WHEN '475'  THEN 'FAIL_TAPE_Y'
            WHEN '500'  THEN 'FAIL_VISIONX'
            WHEN '510'  THEN 'FAIL_RATIO_VISIONX'
            WHEN '600'  THEN 'FAIL_VISIONY'
            WHEN '610'  THEN 'FAIL_RATIO_VISIONY'
            WHEN '700'  THEN 'FAIL_TAPE_EXPIRE'
            WHEN '750'  THEN 'FAIL_TAPE_RETRY3TIMES'
            WHEN '800'  THEN 'FAIL_SILICONE_EXPIRE'
            WHEN '850'  THEN 'SILICONE_BUTTON_PEEL_TEST_SAMPLE'
            ELSE NULL
        END
    FROM ODS.mfg.ProcessHistoryPdrPartProducedEvent
    WHERE TimeStamp >= @Start
      AND TimeStamp <  @End
      AND SourceLocation LIKE '%FRAME_ASSEMBLY%'
      AND SourceLocation LIKE 'PGT14%'
),
Classified AS (
    SELECT
        TimeStamp,
        IsCounted = CASE WHEN FrameStatus IN ('GOOD_SCREW', 'GOOD_REWORK') OR FrameStatus IN (
                        'FAIL_VISIONX', 'FAIL_RATIO_VISIONX', 'FAIL_VISIONY', 'FAIL_RATIO_VISIONY',
                        'FAIL_SILICONE_EXPIRE') THEN 1 ELSE 0 END,
        IsReject  = CASE WHEN FrameStatus IN (
                        'FAIL_VISIONX', 'FAIL_RATIO_VISIONX', 'FAIL_VISIONY', 'FAIL_RATIO_VISIONY',
                        'FAIL_SILICONE_EXPIRE') THEN 1 ELSE 0 END
    FROM FrameEvents
)
SELECT
    Crew        = [ods].[sas].[getcrew](ods.mfg.fn_plant(), TimeStamp),
    FrameTotal  = SUM(IsCounted),
    FrameReject = SUM(IsReject),
    ScrapRate   = CONVERT(DECIMAL(10,5), SUM(IsReject) * 1.0 / NULLIF(SUM(IsCounted), 0)),
    Target      = @ScrapTargetPct,
    ActionLimit = @ScrapActionLimitPct
FROM Classified
GROUP BY
    [ods].[sas].[getcrew](ods.mfg.fn_plant(), TimeStamp)
ORDER BY
    Crew;
