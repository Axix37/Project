
import argparse
import copy
import logging
import os
import re
import statistics
import subprocess
import sys
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.opc.packuri import PackURI
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt


SCRIPT_DIR = Path(__file__).parent.resolve()

_p = argparse.ArgumentParser(add_help=False)
_p.add_argument("--config", default="config.env")
_args = _p.parse_known_args()[0]
load_dotenv(SCRIPT_DIR / _args.config, override=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger(__name__)
log.info("===== Weekly Ops Report — config: %s =====", _args.config)

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

def _server_block(key: str) -> dict:
    return dict(
        server=os.getenv(f"SQL_SERVER_{key}", ""),
        database=os.getenv(f"SQL_DATABASE_{key}", "ODS"),
        username=os.getenv(f"SQL_USERNAME_{key}", ""),
        password=os.getenv(f"SQL_PASSWORD_{key}", ""),
        driver=os.getenv(f"SQL_DRIVER_{key}", "ODBC Driver 17 for SQL Server"),
    )

SERVERS = {"PGT1": _server_block("PGT1"), "PGT2": _server_block("PGT2")}

_PLANT_PREFIX_RE = re.compile(r"^[A-Za-z0-9]+$")
_STATION_RE      = re.compile(r"^[A-Z0-9_]+$")


def parse_plants() -> list[dict]:
    """Parse PLANTS=<prefix>:<label>:<daily_fg_goal>:<server_key>,... from config."""
    raw = os.getenv("PLANTS", "")
    plants = []
    for entry in filter(None, (e.strip() for e in raw.split(","))):
        prefix, label, goal, server_key = (x.strip() for x in entry.split(":"))
        if not _PLANT_PREFIX_RE.match(prefix):
            raise ValueError(f"Invalid PLANTS prefix: {prefix!r}")
        if server_key not in SERVERS:
            raise ValueError(f"Invalid PLANTS server key: {server_key!r}")
        plants.append(dict(prefix=prefix, label=label,
                            daily_goal=float(goal), server_key=server_key))
    return plants


def _parse_pairs(raw: str) -> dict[str, str]:
    """'A=x,B=y' -> {'A': 'x', 'B': 'y'}; order is kept."""
    out = {}
    for item in filter(None, (i.strip() for i in raw.split(","))):
        key, _, val = item.partition("=")
        out[key.strip()] = val.strip()
    return out


def _plant_setting(name: str, plant: dict, default: str = "") -> str:
    """<NAME>_<PREFIX> wins over <NAME>, e.g. AREA_STATIONS_PGT2 over AREA_STATIONS."""
    return os.getenv(f"{name}_{plant['prefix']}", os.getenv(name, default))


def area_stations(plant: dict, setting: str = "AREA_STATIONS") -> dict[str, list[str]]:
    """Area name -> station names, from AREA_STATIONS (per-plant override allowed)."""
    areas = {}
    for area, stations in _parse_pairs(_plant_setting(setting, plant)).items():
        names = [s.strip().upper() for s in stations.split("|") if s.strip()]
        for s in names:
            if not _STATION_RE.match(s):
                raise ValueError(f"Invalid station name in AREA_STATIONS: {s!r}")
        areas[area] = names
    return areas


def ma_stations(plant: dict) -> dict[str, list[str]]:
    """Stations each area's MA is measured on: AREA_MA_STATIONS, else AREA_STATIONS.

    PGT 1.1's frame MA is on the FRAME_C1-C4 cells, while its throughput is
    on the FRAME roll-up (same split as hourly_ma_query.sql).
    """
    return {**area_stations(plant), **area_stations(plant, "AREA_MA_STATIONS")}


def _goal_map(name: str, plant: dict) -> dict[str, float | None]:
    out = {}
    for key, val in _parse_pairs(_plant_setting(name, plant)).items():
        try:
            out[key] = float(val) if val else None
        except ValueError:
            log.warning("%s_%s: %s=%r is not a number — ignored.", name, plant["prefix"], key, val)
            out[key] = None
    return out


PLANTS = parse_plants()

# every value embedded directly into a SQL string must appear here.
_ALLOWED_PLANT_FILTERS = {f"{p['prefix']}%" for p in PLANTS}

SHIFT_START_HOUR = int(os.getenv("SHIFT_START_HOUR", "6"))
if not 0 <= SHIFT_START_HOUR <= 23:
    raise ValueError(f"SHIFT_START_HOUR must be 0-23, got {SHIFT_START_HOUR}")
# Throughput numbers (FG per day, Critical Area, constraint, tool loss) count
# days from this hour; only the crew metrics use SHIFT_START_HOUR.
DAY_START_HOUR = int(os.getenv("THROUGHPUT_DAY_START_HOUR", "0"))
if not 0 <= DAY_START_HOUR <= 23:
    raise ValueError(f"THROUGHPUT_DAY_START_HOUR must be 0-23, got {DAY_START_HOUR}")

CREWS                = [c.strip() for c in os.getenv("CREWS", "A,B,C,D").split(",") if c.strip()]
CREW_THROUGHPUT_AREA = os.getenv("CREW_THROUGHPUT_AREA", "Framing")

CREW_WINDOW_DAYS     = int(os.getenv("CREW_WINDOW_DAYS", "14"))
# "total": each crew's throughput over the crew window; "per_shift": that
# total divided by the shift days the crew worked.
CREW_THROUGHPUT_MODE = os.getenv("CREW_THROUGHPUT_MODE", "total").strip().lower()
if CREW_THROUGHPUT_MODE not in ("total", "per_shift"):
    raise ValueError(f"CREW_THROUGHPUT_MODE must be total or per_shift, got {CREW_THROUGHPUT_MODE!r}")
if CREW_WINDOW_DAYS < 1:
    raise ValueError(f"CREW_WINDOW_DAYS must be at least 1, got {CREW_WINDOW_DAYS}")

QUERY_FILES = dict(
    area_hourly=os.getenv("AREA_HOURLY_QUERY_FILE", ""),
    fg_daily=os.getenv("FG_DAILY_QUERY_FILE", ""),
    crew_fg=os.getenv("CREW_SCORECARD_FG_QUERY_FILE", ""),
)
REWORK_RERUN_WORKBOOK = os.getenv("REWORK_RERUN_WORKBOOK", "")
REWORK_RERUN_SHEET    = os.getenv("REWORK_RERUN_SHEET", "") or 0
REFRESH_WORKBOOK      = os.getenv("REFRESH_WORKBOOK", "true").lower() in ("1", "true", "yes")


TOOL_LOSS_FILE           = os.getenv("TOOL_LOSS_FILE", "")
TOOL_LOSS_SHEET          = os.getenv("TOOL_LOSS_SHEET", "") or 0
TOOL_LOSS_COLUMNS        = {"Date": "Date", "Plant": "Plant", "Tool": "Tool", "Area": "",
                            "ModulesLost": "Modules Lost", "Downtime": "Downtime",
                            **_parse_pairs(os.getenv("TOOL_LOSS_COLUMNS", ""))}
TOOL_LOSS_DOWNTIME_UNIT  = os.getenv("TOOL_LOSS_DOWNTIME_UNIT", "min")
TOOL_LOSS_TOP_N          = int(os.getenv("TOOL_LOSS_TOP_N", "3"))
TOOL_LOSS_WATCH_INCREASE = float(os.getenv("TOOL_LOSS_WATCH_INCREASE", "0.5"))
TOOL_LOSS_SUMMARY_FILE   = os.getenv("TOOL_LOSS_SUMMARY_FILE", "")


CONSTRAINT_QUERY_FILE = os.getenv("CONSTRAINT_QUERY_FILE", "")
CONSTRAINT_EXCLUDE    = [x.strip().upper() for x in
                         os.getenv("CONSTRAINT_EXCLUDE", "REWORK|CVS|NPR|LED|SCRIBE").split("|") if x.strip()]
CONSTRAINT_IDLE_MAX   = float(os.getenv("CONSTRAINT_IDLE_MAX", "0.15"))
CONSTRAINT_CT_SOURCE  = os.getenv("CONSTRAINT_CT_SOURCE", "actual").strip().lower()
CONSTRAINT_MIN_CYCLES = int(os.getenv("CONSTRAINT_MIN_CYCLES", "5"))
CONSTRAINT_ACTUAL_CT_RANGE = tuple(float(x) for x in os.getenv("CONSTRAINT_ACTUAL_CT_RANGE", "0.5,3").split(","))
if CONSTRAINT_CT_SOURCE not in ("actual", "target"):
    raise ValueError(f"CONSTRAINT_CT_SOURCE must be actual or target, got {CONSTRAINT_CT_SOURCE!r}")
if len(CONSTRAINT_ACTUAL_CT_RANGE) != 2:
    raise ValueError("CONSTRAINT_ACTUAL_CT_RANGE must be two numbers, e.g. 0.5,3")

.
AREA_LOSS_THRESHOLD = float(os.getenv("AREA_LOSS_THRESHOLD", "0.02"))
AREA_LOSS_TOP_N     = int(os.getenv("AREA_LOSS_TOP_N", "3"))

# Cells are colored against their goal: green meets it, yellow misses it by
# up to GOAL_WARN_BAND (a fraction of the goal), red misses it by more.
GOAL_WARN_BAND = float(os.getenv("GOAL_WARN_BAND", "0.05"))

# Target line on the Throughput FG chart: modules per day for all plants.

FG_DAILY_TARGET = os.getenv("FG_DAILY_TARGET", "")
# Top of the FG per day chart's y axis; blank lets PowerPoint pick it.
FG_CHART_MAX = os.getenv("FG_CHART_MAX", "23000")

REPORT_WEEK_ENDING = os.getenv("REPORT_WEEK_ENDING", "").strip()

PPTX_TEMPLATE   = os.getenv("PPTX_TEMPLATE", "Template.pptx")
PPTX_OUTPUT_DIR = os.getenv("PPTX_OUTPUT_DIR", "")

log.info("Plants: %s", ", ".join(f"{p['label']} ({p['prefix']}%)" for p in PLANTS))

NOT_CONFIGURED = "n/a"   # source not set up yet
NO_GOAL        = "–"     # goal not set in config.env


def report_end(week_ending: str = "", today: date | None = None) -> date:
    """The Monday after the week containing week_ending; blank = last week."""
    day = date.fromisoformat(week_ending) + timedelta(days=7) if week_ending else (today or date.today())
    return day - timedelta(days=day.weekday())


def report_window(end: date | None = None, days: int = 7) -> list[date]:
    end = end or REPORT_END
    return [end - timedelta(days=d) for d in range(days, 0, -1)]


REPORT_END  = report_end(REPORT_WEEK_ENDING)
WINDOW      = report_window()
CREW_WINDOW = report_window(days=CREW_WINDOW_DAYS)

# ---------------------------------------------------------------------------
# SQL helpers
# ---------------------------------------------------------------------------

def _conn_string(server_key: str) -> str:
    s = SERVERS[server_key]
    base = f"DRIVER={{{s['driver']}}};SERVER={s['server']};DATABASE={s['database']};"
    if s["username"] and s["password"]:
        return base + f"UID={s['username']};PWD={s['password']};"
    return base + "Trusted_Connection=yes;"


def _run_query(server_key: str, sql: str) -> pd.DataFrame:
    import pyodbc 

    conn = pyodbc.connect(_conn_string(server_key), timeout=60)
    try:
        cur = conn.cursor()
        cur.execute(sql)
        while cur.description is None:
            if not cur.nextset():
                return pd.DataFrame()
        cols = [c[0] for c in cur.description]
        rows = cur.fetchall()
        return pd.DataFrame.from_records(rows, columns=cols)
    finally:
        conn.close()


def _load_sql(query_file: str) -> str:
    path = (SCRIPT_DIR / query_file).resolve()
    return path.read_text(encoding="utf-8")


def _plant_filter(plant: dict) -> str:
    pf = f"{plant['prefix']}%"
    if pf not in _ALLOWED_PLANT_FILTERS:
        raise ValueError(f"Plant filter '{pf}' not in allowlist.")
    return pf


def _set_report_end(sql: str) -> str:
    """End the query's window at REPORT_END instead of today."""
    try:
        return _set_declare(sql, "ReportEnd", f"'{REPORT_END:%Y-%m-%d}'")
    except ValueError:
        log.warning("A query has no @ReportEnd (an older copy?), so it counts back from today, "
                    "not from %s: %s", REPORT_END, sql.strip().splitlines()[1][:70])
        return sql


def _fill_sql(sql: str, plant: dict, window_days: int = 7, start_hour: int | None = None,
              **extra: str) -> str:
    """Fill a query's placeholders; {shift_start_hour} is start_hour when given."""
    hour = SHIFT_START_HOUR if start_hour is None else start_hour
    sql = (_set_report_end(sql).replace("{plant_filter}", _plant_filter(plant))
              .replace("{shift_start_hour}", str(hour))
              .replace("{day_start_hour}", str(DAY_START_HOUR))
              .replace("{window_days}", str(int(window_days))))
    for key, val in extra.items():
        sql = sql.replace("{" + key + "}", val)
    return sql


def _safe_query(label: str, plant: dict, sql_fn) -> pd.DataFrame | None:
    """Run one query for one plant; a failure only blanks that section."""
    if not SERVERS[plant["server_key"]]["server"]:
        return None
    try:
        return _run_query(plant["server_key"], sql_fn())
    except Exception as exc:  # noqa: BLE001 — one bad source must not stop the deck
        log.error("%s failed for %s: %s", label, plant["label"], exc)
        return None

# ---------------------------------------------------------------------------
# Fetchers — each returns plain DataFrames/dicts, or None when not available
# ---------------------------------------------------------------------------

def fetch_station_hours(plant: dict) -> pd.DataFrame | None:
    """Per equipment x shift day x crew run time, active tools only, with Area.

    Covers the longer of the deck and crew windows; area_metrics narrows it
    back to WINDOW by ReadDate (days from DAY_START_HOUR), crew_throughput
    to CREW_WINDOW by ShiftDate (days from SHIFT_START_HOUR).
    """
    if not QUERY_FILES["area_hourly"]:
        return None
    station_to_area = {s: a for a, stations in area_stations(plant).items() for s in stations}
    station_to_ma_area = {s: a for a, stations in ma_stations(plant).items() for s in stations}
    if not station_to_area:
        return None
    station_list = ",".join(f"'{s}'" for s in {**station_to_area, **station_to_ma_area})
    df = _safe_query("Station run-time query", plant, lambda: _fill_sql(
        _load_sql(QUERY_FILES["area_hourly"]), plant,
        window_days=max(len(WINDOW), len(CREW_WINDOW)), station_list=station_list))
    if df is None:
        return None
    if df.empty:  # the query ran but the line logged nothing this week
        return pd.DataFrame(columns=["Equipment", "Station", "ShiftDate", "ReadDate", "Crew",
                                     "UpTimeCorrected", "MaxMinutes", "GoodModule", "Area", "MAArea"])
    for col in ("UpTimeCorrected", "MaxMinutes", "GoodModule"):
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    df["ShiftDate"] = pd.to_datetime(df["ShiftDate"]).dt.date
    if "ReadDate" in df.columns:
        df["ReadDate"] = pd.to_datetime(df["ReadDate"]).dt.date
    else:
        log.warning("%s: the station run-time query has no ReadDate column (an older copy of "
                    "weekly_area_hourly_query.sql?) — Critical Area days use the crew shift day.",
                    plant["label"])
        df["ReadDate"] = df["ShiftDate"]
    df["Area"] = df["Station"].str.upper().map(station_to_area)
    df["MAArea"] = df["Station"].str.upper().map(station_to_ma_area)
    # Tools that never ran in the window (decommissioned, idle spares) would
    # only inflate the MA denominator.
    active = df.groupby("Equipment")[["UpTimeCorrected", "GoodModule"]].sum()
    active = active[(active["UpTimeCorrected"] > 0) | (active["GoodModule"] > 0)].index
    return df[df["Equipment"].isin(active)]


def fetch_fg_daily(plant: dict) -> dict[date, float] | None:
    if not QUERY_FILES["fg_daily"]:
        return None
    df = _safe_query("FG daily query", plant,
                     lambda: _fill_sql(_load_sql(QUERY_FILES["fg_daily"]), plant,
                                       start_hour=DAY_START_HOUR))
    if df is None:
        return None
    qty = pd.to_numeric(df["Qty"], errors="coerce").fillna(0.0) if not df.empty else []
    by_day = {pd.to_datetime(d).date(): float(q) for d, q in zip(df.get("ShiftDate", []), qty)}
    return {d: by_day.get(d, 0.0) for d in WINDOW}


def fetch_crew_fg(plant: dict) -> dict[str, float] | None:
    if not QUERY_FILES["crew_fg"]:
        return None
    df = _safe_query("Crew FG query", plant,
                     lambda: _fill_sql(_load_sql(QUERY_FILES["crew_fg"]), plant,
                                       window_days=len(CREW_WINDOW)))
    if df is None:
        return None
    return {str(r["Crew"]): float(r["GoodModule"] or 0) for _, r in df.iterrows() if r["Crew"]}


def _set_declare(sql: str, name: str, value: str) -> str:
    """Replace the value of `DECLARE @<name> <type> = ...;` in a query file."""
    pattern = re.compile(rf"(DECLARE\s+@{name}\s+[A-Z0-9_(), ]+?=\s*)[^;]+;", re.IGNORECASE)
    if not pattern.search(sql):
        raise ValueError(f"query has no '@{name}' declaration to set")
    return pattern.sub(lambda m: f"{m.group(1)}{value};", sql, count=1)


def _crew_window_sql(query_file: str, plant: dict, plant_var: bool = False) -> str:
    """Load a crew_*_14d.sql query and set its window (and @Plant) for this run."""
    sql = _set_report_end(_load_sql(query_file))
    sql = _set_declare(sql, "ShiftStartHour", str(SHIFT_START_HOUR))
    sql = _set_declare(sql, "WindowDays", str(len(CREW_WINDOW)))
    if plant_var:
        sql = _set_declare(sql, "Plant", f"'{_plant_filter(plant)}'")
    return sql


def fetch_frame_scrap(plant: dict) -> dict[str, float] | None:
    """Frame scrap % by crew over the crew window.

    Runs the plant's crew_frame_scrap_14d_*.sql (each plant logs frame
    scrap in a different table), with its window set to the crew window,
    and takes rejects / totals per crew.
    """
    query_file = _plant_setting("FRAME_SCRAP_QUERY_FILE", plant)
    if not query_file:
        return None
    df = _safe_query("Frame scrap query", plant, lambda: _crew_window_sql(query_file, plant))
    if df is None or df.empty:
        return None
    total_col = next((c for c in df.columns if c.endswith("Total")), None)
    reject_col = next((c for c in df.columns if c.endswith("Reject")), None)
    if not total_col or not reject_col:
        log.error("Frame scrap query for %s has no *Total / *Reject columns.", plant["label"])
        return None
    sums = df.groupby("Crew")[[total_col, reject_col]].sum()
    return {str(crew): (float(r[reject_col]) / float(r[total_col]))
            for crew, r in sums.iterrows() if r[total_col]}


def fetch_catchall(plant: dict) -> dict[str, float] | None:
    """Catchall by crew, from crew_catchall_14d.sql (returns Crew, Catchall).

    Its @Plant, @WindowDays and @ShiftStartHour are set for each plant and
    the crew window.
    """
    query_file = _plant_setting("CATCHALL_QUERY_FILE", plant)
    if not query_file:
        return None
    df = _safe_query("Catchall query", plant,
                     lambda: _crew_window_sql(query_file, plant, plant_var=True))
    if df is None:
        return None
    col = "Catchall" if "Catchall" in df.columns else "Value"
    if col not in df.columns:
        log.error("Catchall query for %s has no Catchall column.", plant["label"])
        return None
    return {str(r["Crew"]): float(r[col]) for _, r in df.iterrows() if pd.notna(r[col])}


def refresh_workbook(path: Path) -> None:
    """Refresh every Power BI connection in the workbook through Excel (Windows only)."""
    if os.name != "nt" or not REFRESH_WORKBOOK:
        return
    script = SCRIPT_DIR / "refresh_workbook.ps1"
    log.info("Refreshing %s through Excel ...", path.name)
    result = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script), "-Path", str(path)],
        capture_output=True, text=True, timeout=900)
    if result.returncode != 0:
        log.error("Workbook refresh failed — using the last saved values. %s", result.stderr.strip())


def fetch_rework_rerun(plant: dict) -> dict[str, float] | None:
    """Rework rerun % by crew, from crew_rework_rerun_14d.sql (Crew, ReworkRerun).

    The Hipot rerun tables live in MesReporting on PGT1MESSQLODS for every
    site, so the query runs on REWORK_RERUN_SERVER (PGT1 unless set per
    plant). Its @Plant, @WindowDays and @ShiftStartHour are set per run.
    """
    query_file = _plant_setting("REWORK_RERUN_QUERY_FILE", plant)
    if not query_file:
        return None
    server_key = _plant_setting("REWORK_RERUN_SERVER", plant, "PGT1")
    if server_key not in SERVERS:
        log.error("REWORK_RERUN_SERVER=%r is not a server key (PGT1 / PGT2).", server_key)
        return None
    df = _safe_query("Rework rerun query", {**plant, "server_key": server_key},
                     lambda: _crew_window_sql(query_file, plant, plant_var=True))
    if df is None:
        return None
    if "ReworkRerun" not in df.columns:
        log.error("Rework rerun query for %s has no ReworkRerun column.", plant["label"])
        return None
    return {str(r["Crew"]): float(r["ReworkRerun"]) for _, r in df.iterrows()
            if r["Crew"] and pd.notna(r["ReworkRerun"])}


def load_rework_rerun() -> pd.DataFrame | None:
    """Rework rerun by plant and crew, from the Power BI-connected workbook.

    Expects a sheet with a header row holding Plant, Crew and ReworkRerun
    columns (a PivotTable in tabular layout works). Plant must match the
    labels in PLANTS, e.g. 'PGT 1.1'.
    """
    if not REWORK_RERUN_WORKBOOK:
        return None
    path = (SCRIPT_DIR / REWORK_RERUN_WORKBOOK).resolve()
    if not path.exists():
        log.warning("Rework rerun workbook not found: %s", path)
        return None
    refresh_workbook(path)
    try:
        raw = pd.read_excel(path, sheet_name=REWORK_RERUN_SHEET, header=None)
    except Exception as exc:  # noqa: BLE001
        log.error("Could not read %s: %s", path, exc)
        return None
    wanted = {"plant", "crew", "reworkrerun"}
    for i, row in raw.iterrows():
        names = [str(v).strip().replace(" ", "").lower() for v in row.values]
        if wanted.issubset(names):
            df = raw.iloc[i + 1:].copy()
            df.columns = names
            df = df.dropna(subset=["plant", "crew"])
            return pd.DataFrame(dict(Plant=df["plant"].astype(str).str.strip(),
                                     Crew=df["crew"].astype(str).str.strip(),
                                     Value=pd.to_numeric(df["reworkrerun"], errors="coerce")))
    log.error("%s: no header row with Plant, Crew and ReworkRerun columns.", path.name)
    return None


TOOL_LOSS_METRICS = ("ModulesLost", "Downtime")


def _plant_key(name) -> str:
    """'PGT 1.1', 'PGT11' and 'pgt11' all match the same plant."""
    return re.sub(r"[^a-z0-9]", "", str(name).lower())


def load_tool_loss() -> pd.DataFrame | None:
    """Tool loss log rows: ShiftDate, Plant (PLANTS label), Tool and the metrics.

    Reads TOOL_LOSS_FILE (.xlsx or .csv). The header row is found by the
    column names in TOOL_LOSS_COLUMNS, so title rows above it are fine. A
    metric whose column name is left blank is not reported. Timestamps with
    a time of day count on the day that starts at THROUGHPUT_DAY_START_HOUR;
    plain dates count as they are.
    """
    if not TOOL_LOSS_FILE:
        return None
    path = (SCRIPT_DIR / TOOL_LOSS_FILE).resolve()
    if not path.exists():
        log.warning("Tool loss file not found: %s", path)
        return None
    cols = {k: v.strip().replace(" ", "").lower() for k, v in TOOL_LOSS_COLUMNS.items() if v.strip()}
    metrics = [m for m in TOOL_LOSS_METRICS if m in cols]
    if not {"Date", "Plant", "Tool"} <= cols.keys() or not metrics:
        log.error("TOOL_LOSS_COLUMNS needs Date, Plant, Tool and ModulesLost and/or Downtime.")
        return None
    try:
        if path.suffix.lower() == ".csv":
            raw = pd.read_csv(path, header=None, dtype=object)
        else:
            raw = pd.read_excel(path, sheet_name=TOOL_LOSS_SHEET, header=None)
    except Exception as exc:  # noqa: BLE001
        log.error("Could not read %s: %s", path, exc)
        return None
    for i, row in raw.iterrows():
        names = [str(v).strip().replace(" ", "").lower() for v in row.values]
        if set(cols.values()) <= set(names):
            break
    else:
        log.error("%s: no header row with columns %s.", path.name,
                  ", ".join(TOOL_LOSS_COLUMNS[k] for k in cols))
        return None
    body = raw.iloc[i + 1:]
    col = {k: body.iloc[:, names.index(v)] for k, v in cols.items()}

    ts = pd.to_datetime(col["Date"], errors="coerce")
    has_time = ts.notna() & (ts != ts.dt.normalize())
    ts = ts.where(~has_time, ts - pd.Timedelta(hours=DAY_START_HOUR))
    labels = {}
    for p in PLANTS:
        labels[_plant_key(p["label"])] = labels[_plant_key(p["prefix"])] = p["label"]
    plant = col["Plant"].map(lambda v: labels.get(_plant_key(v)))
    unknown = set(col["Plant"][plant.isna() & col["Plant"].notna()].astype(str))
    if unknown:
        log.warning("Tool loss: plants not in PLANTS, skipped: %s", ", ".join(sorted(unknown)))

    df = pd.DataFrame(dict(ShiftDate=ts.dt.date, Plant=plant,
                           Tool=col["Tool"].astype(str).str.strip()))
    if "Area" in col:
        df["Area"] = col["Area"].where(col["Area"].notna(), "").astype(str).str.strip()
    for m in metrics:
        df[m] = pd.to_numeric(col[m], errors="coerce").fillna(0.0)
    df = df.dropna(subset=["ShiftDate", "Plant"])
    return df[df["Tool"].ne("") & df["Tool"].ne("nan")]


def fetch_constraint_data(plant: dict) -> pd.DataFrame | None:
    """Per tool x hour: uptime, blocked, starved, output, CT target and actual CT.

    From weekly_constraint_query.sql, every tool with a CT target.
    """
    if not CONSTRAINT_QUERY_FILE:
        return None
    df = _safe_query("Constraint query", plant, lambda: _fill_sql(
        _load_sql(CONSTRAINT_QUERY_FILE), plant, window_days=len(WINDOW),
        start_hour=DAY_START_HOUR))
    if df is None:
        return None
    for col in ("MaxMinutes", "UpTimeCorrected", "Blocked", "Starved", "Output", "Cycles"):
        df[col] = pd.to_numeric(df.get(col), errors="coerce").fillna(0.0)
    for col in ("CTTarget", "ActualCT"):
        df[col] = pd.to_numeric(df.get(col), errors="coerce")
    df["ShiftDate"] = pd.to_datetime(df["ShiftDate"]).dt.date
    df["ReadTimeHour"] = pd.to_datetime(df["ReadTimeHour"])
    df["Station"] = df["Station"].astype(str).str.upper()
    return df

# ---------------------------------------------------------------------------
# Metrics
# ---------------------------------------------------------------------------

def area_metrics(hours: pd.DataFrame | None, plant: dict) -> dict[str, dict]:
    """Per area: MA, daily throughput avg / median, week total, utilization."""
    goals = _goal_map("AREA_WEEKLY_GOALS", plant)
    out = {}
    for area in area_stations(plant):
        if hours is None:
            out[area] = None
            continue
        rows = hours[(hours["Area"] == area) & hours["ReadDate"].isin(WINDOW)]
        ma_rows = hours[(hours["MAArea"] == area) & hours["ReadDate"].isin(WINDOW)]
        max_min = ma_rows["MaxMinutes"].sum()
        daily = rows.groupby("ReadDate")["GoodModule"].sum().to_dict()
        per_day = [float(daily.get(d, 0.0)) for d in WINDOW]
        total = sum(per_day)
        goal = goals.get(area)
        out[area] = dict(
            ma=(ma_rows["UpTimeCorrected"].sum() / max_min) if max_min else None,
            avg=statistics.mean(per_day),
            median=statistics.median(per_day),
            total=total,
            utilization=(total / goal) if goal else None,
        )
    return out


def crew_throughput(hours: pd.DataFrame | None) -> dict[str, float] | None:
    """Throughput at CREW_THROUGHPUT_AREA over CREW_WINDOW, per crew.

    With CREW_THROUGHPUT_MODE=per_shift it is divided by the distinct shift
    days the crew worked.
    """
    if hours is None:
        return None
    rows = hours[(hours["Area"] == CREW_THROUGHPUT_AREA) & hours["Crew"].notna()
                 & hours["ShiftDate"].isin(CREW_WINDOW)]
    out = {}
    for crew, grp in rows.groupby("Crew"):
        shifts = grp["ShiftDate"].nunique()
        if not shifts:
            continue
        total = grp["GoodModule"].sum()
        out[str(crew)] = total / shifts if CREW_THROUGHPUT_MODE == "per_shift" else total
    return out


def constraint_process_names(plant: dict) -> tuple[dict[str, str], set[str]]:
    """Station -> process name for the constraint, and roll-up stations to skip.

    Stations in an area (AREA_STATIONS, with AREA_MA_STATIONS cells winning)
    take the area's name, so PGT 1.1's four frame cells count as one
    Framing process. A station an MA override replaced (PGT 1.1's FRAME
    roll-up) is skipped so its modules aren't counted twice. Other
    stations are grouped in daily_constraints by name.
    """
    groups = ma_stations(plant)
    names = {st: area for area, stations in groups.items() for st in stations}
    rollups = {st for stations in area_stations(plant).values() for st in stations} - set(names)
    return names, rollups


_CELL_SUFFIX_RE = r"_C?\d+$"   # FRAME_C1-> FRAME
_TOOL_NAME_RE   = re.compile(r"^[A-Z0-9_-]+$")


def constraint_tool_list(plant: dict) -> dict[str, str]:
    """CONSTRAINT_CANDIDATES(_<prefix>): tool name -> process, when the list is set.

    `Process=NAME|NAME,...`, where NAME is a station (LAMINATOR, every
    unit's laminator) or one piece of equipment (PGT14A-LAMINATOR). The
    names under one process are added up as that process's capacity.
    Empty dict = no list, every process with a CT competes.
    """
    out = {}
    for proc, tools in _parse_pairs(_plant_setting("CONSTRAINT_CANDIDATES", plant)).items():
        for name in (t.strip().upper() for t in tools.split("|") if t.strip()):
            if not _TOOL_NAME_RE.match(name):
                raise ValueError(f"Invalid tool name in CONSTRAINT_CANDIDATES: {name!r}")
            out[name] = proc
    return out


def _rate_factors(plant: dict) -> dict[str, float]:
    """CONSTRAINT_RATE_FACTOR(_<prefix>): STATION=modules per CT, default 1."""
    out = {}
    for station, val in _parse_pairs(_plant_setting("CONSTRAINT_RATE_FACTOR", plant)).items():
        try:
            out[station.strip().upper()] = float(val)
        except ValueError:
            log.warning("CONSTRAINT_RATE_FACTOR: %s=%r is not a number — ignored.", station, val)
    return out


def _capacity(df: pd.DataFrame, plant: dict) -> pd.DataFrame:
    """Add CTUsed, ActualCTUsed and IdealModules (uptime hours x 3600 / CT) per tool-hour.

    The tool's actual CT that hour when CONSTRAINT_CT_SOURCE=actual and it
    ran at least CONSTRAINT_MIN_CYCLES cycles at an actual CT inside
    CONSTRAINT_ACTUAL_CT_RANGE x its target; otherwise its CT target, times
    CONSTRAINT_RATE_FACTOR (actual CT is already per module).
    """
    low, high = CONSTRAINT_ACTUAL_CT_RANGE
    target = df["CTTarget"]
    actual_ok = (df["ActualCT"].notna() & (df["Cycles"] >= CONSTRAINT_MIN_CYCLES)
                 & df["ActualCT"].between(low * target, high * target))
    if CONSTRAINT_CT_SOURCE == "target":
        actual_ok &= False
    factor = df["Station"].map(_rate_factors(plant)).fillna(1.0)
    ct = df["ActualCT"].where(actual_ok, target / factor)
    df = df.assign(CTUsed=ct, ActualCTUsed=actual_ok,
                   IdealModules=(df["UpTimeCorrected"] / 60 * 3600 / ct).where(ct > 0, 0.0))
    if CONSTRAINT_CT_SOURCE == "actual":
        ran = df[df["UpTimeCorrected"] > 0]
        share = ran.groupby("Process")["ActualCTUsed"].mean()
        for proc, used in share[share < 0.5].items():
            log.info("Constraint %s: %s used its CT target for %.0f%% of its running hours "
                     "(too few cycles or actual CT out of range).", plant["label"], proc, 100 * (1 - used))
    return df


def daily_constraints(df: pd.DataFrame | None, plant: dict) -> dict[date, dict] | None:
    """Each hour's constraint, summed up per shift day.

    capacity = sum of uptime hours x 3600 / CT over the process's tools
    that ran that hour (CT: see _capacity). Processes blocked + starved for
    more than CONSTRAINT_IDLE_MAX of their uptime are passed over; when all
    were, the lowest is kept unconfirmed. Each day is named for the process
    that was the constraint for the most hours. Days with no data are left out.

    With CONSTRAINT_CANDIDATES set for the plant, only the listed tools compete,
    grouped exactly as listed. Without it, every process with a CT does,
    grouped by area and by name.
    """
    if df is None:
        return None
    df = constraint_processes(df, plant)
    return _pick_constraints(_capacity(df[(df["UpTimeCorrected"] > 0) | (df["Output"] > 0)], plant), plant)


def constraint_processes(df: pd.DataFrame, plant: dict) -> pd.DataFrame:
    """The candidate tools' rows, each with its Process.

    CONSTRAINT_CANDIDATES when set; otherwise every tool with a CT, grouped
    by area and by name. Hours a tool was down are kept (downtime is a loss).
    """
    listed = constraint_tool_list(plant)
    if listed:
        return _listed_tools(df, listed, plant)
    names, rollups = constraint_process_names(plant)
    keep = ~df["Station"].isin(rollups)
    for pattern in CONSTRAINT_EXCLUDE:
        keep &= ~df["Station"].str.contains(pattern, regex=False)
    df = df[keep].copy()
    if df.empty:
        return df.assign(Process=pd.Series(dtype=object))
  
    base = df["Station"].str.replace(_CELL_SUFFIX_RE, "", regex=True)
    cell_groups = set(base[base != df["Station"]])
    rollup = ~df["Station"].isin(names) & df["Station"].isin(cell_groups)
    if rollup.any():
        log.info("Constraint %s: skipped roll-up stations %s", plant["label"],
                 ", ".join(sorted(set(df.loc[rollup, "Station"]))))
    df, base = df[~rollup], base[~rollup]
    df["Process"] = df["Station"].map(names).fillna(base.str.replace("_", " ").str.title())
    return df


def _listed_tools(df: pd.DataFrame, listed: dict[str, str], plant: dict) -> pd.DataFrame:
    """Only the CONSTRAINT_CANDIDATES rows, with their listed process.

    An equipment name wins over its station name, so one unit can be listed
    on its own.
    """
    equipment = df["Equipment"].astype(str).str.upper()
    process = equipment.map(listed).fillna(df["Station"].map(listed))
    missing = set(listed) - set(equipment) - set(df["Station"])
    if missing:
        log.warning("Constraint %s: CONSTRAINT_CANDIDATES names with no run-time rows this week: %s",
                    plant["label"], ", ".join(sorted(missing)))
    df = df[process.notna()].copy()
    df["Process"] = process[df.index]
    return df


def _pick_constraints(df: pd.DataFrame, plant: dict) -> dict[date, dict]:
    """Each hour's constraint, then per shift day the process that held it most hours."""
    if df.empty:
        return {}
    df = df.assign(Idle=df["Blocked"] + df["Starved"], CTxUp=df["CTUsed"] * df["UpTimeCorrected"],
                   TargetxUp=df["CTTarget"] * df["UpTimeCorrected"])
    sums = dict(Capacity=("IdealModules", "sum"), Uptime=("UpTimeCorrected", "sum"),
                MaxMinutes=("MaxMinutes", "sum"), Idle=("Idle", "sum"), Output=("Output", "sum"))
    hourly = df.groupby(["ShiftDate", "ReadTimeHour", "Process"]).agg(**sums)
    hourly = hourly[hourly["Capacity"] > 0]
    daily = df.groupby(["ShiftDate", "Process"]).agg(
        Tools=("Equipment", "nunique"), CTxUp=("CTxUp", "sum"), TargetxUp=("TargetxUp", "sum"), **sums)

  
    week = daily.groupby("Process")[["Output", "Capacity"]].sum()
    week = week[week["Capacity"] > 0]
    for proc, r in week[week["Output"] > 1.15 * week["Capacity"]].iterrows():
        log.warning("Constraint %s: %s made %.0f%% of its capacity at CT — check its CT, or set "
                    "CONSTRAINT_RATE_FACTOR for it.", plant["label"], proc, 100 * r["Output"] / r["Capacity"])

    for proc, r in week[week["Output"] < 0.4 * week["Capacity"]].iterrows():
        log.warning("Constraint %s: %s made only %.0f%% of its capacity at CT this week — if it ran "
                    "normally, its CT may not be per module; check CONSTRAINT_RATE_FACTOR.",
                    plant["label"], proc, 100 * r["Output"] / r["Capacity"])

  
    hourly["IdleShare"] = (hourly["Idle"] / hourly["Uptime"]).where(hourly["Uptime"] > 0)
    picks = []
    for (day, hour), grp in hourly.groupby(level=["ShiftDate", "ReadTimeHour"]):
        grp = grp.droplevel(["ShiftDate", "ReadTimeHour"]).sort_values("Capacity")
        ok = grp["IdleShare"].le(CONSTRAINT_IDLE_MAX)
        pos = int(ok.to_numpy().argmax()) if ok.any() else 0
        nxt = grp.iloc[pos + 1] if pos + 1 < len(grp) else None
        picks.append(dict(
            day=day, hour=hour, process=grp.index[pos], capacity=float(grp.iloc[pos]["Capacity"]),
            confirmed=bool(ok.iloc[pos]), skipped=list(grp.index[:pos]),
            next=grp.index[pos + 1] if nxt is not None else None,
            next_capacity=float(nxt["Capacity"]) if nxt is not None else None))

    out = {}
    for day, day_picks in pd.DataFrame(picks).groupby("day"):
        hours = day_picks["process"].value_counts()
        proc = hours.index[0]
        r = daily.loc[(day, proc)]
        with_next = day_picks[day_picks["next"].notna()]
        skipped = pd.Series([p for s in day_picks["skipped"] for p in s], dtype=object).value_counts()
        out[day] = dict(
            process=proc, hours=hours.to_dict(), n_hours=len(day_picks),
           
            capacity=float(day_picks["capacity"].sum()),
            margin=float((with_next["next_capacity"] - with_next["capacity"]).sum()) if len(with_next) else None,
            confirmed=bool(day_picks["confirmed"].mean() >= 0.5),
            tools=int(r["Tools"]), ma=r["Uptime"] / r["MaxMinutes"] if r["MaxMinutes"] else None,
            oe=r["Output"] / r["Capacity"] if r["Capacity"] else None,
            idle=r["Idle"] / r["Uptime"] if r["Uptime"] else None,
            ct=r["CTxUp"] / r["Uptime"] if r["Uptime"] else None,
            ct_target=r["TargetxUp"] / r["Uptime"] if r["Uptime"] else None,
            skipped=skipped.to_dict(), hourly=day_picks.drop(columns="day").to_dict("records"),
            goal=plant["daily_goal"])
    return out


LOSS_BUCKETS = ("Downtime", "CT", "Waiting", "Quality")


def area_losses(df: pd.DataFrame | None, plant: dict) -> dict | None:
    """Per day and candidate process: good modules vs daily goal, and where the shortfall went.

    Per tool and hour, at the CT target rate (3600 / CT target modules an
    hour, with CONSTRAINT_RATE_FACTOR), the modules the tool could have made
    in its scheduled time split into:
        Downtime  scheduled minutes it wasn't up
        Waiting   minutes it was up but blocked or starved (another area's loss)
        CT        running minutes it made fewer parts than CT target allows
                  (slow cycles, micro-stops; negative when it beat the target)
        Quality   parts made that weren't good modules
        Good      good modules
    which add up exactly to its capacity at CT target and 100% MA. When that
    capacity is below the goal, the difference is no tool's loss: the
    process can't make goal even running perfectly. A process is short when
    its good modules miss the goal by more than AREA_LOSS_THRESHOLD of it. Tools that never ran all
    week and processes with no good-module counts are left out.
    """
    if df is None:
        return None
    df = constraint_processes(df, plant)
    active = df.groupby("Equipment")[["UpTimeCorrected", "Output"]].sum()
    active = active[(active["UpTimeCorrected"] > 0) | (active["Output"] > 0)].index
    df = df[df["Equipment"].isin(active) & df["CTTarget"].gt(0)].copy()
    if df.empty:
        return {}
    rate = 60 / (df["CTTarget"] / df["Station"].map(_rate_factors(plant)).fillna(1.0))  # modules a minute
    up = df["UpTimeCorrected"].clip(upper=df["MaxMinutes"])
    idle = (df["Blocked"] + df["Starved"]).clip(upper=up)
    df["Downtime"] = (df["MaxMinutes"] - up) * rate
    df["Waiting"] = idle * rate
    df["CT"] = (up - idle) * rate - df["Output"]
    df["Quality"] = df["Output"] - df["GoodModule"]
    df["Good"] = df["GoodModule"]
    ok = df["ActualCT"].notna() & (df["Cycles"] >= CONSTRAINT_MIN_CYCLES)
    df["ActualxUp"] = df["ActualCT"].where(ok, 0) * up.where(ok, 0)
    df["UpActual"] = up.where(ok, 0)
    df["TargetxUp"] = df["CTTarget"] * up
    df["Up"] = up
    cols = ["Good", *LOSS_BUCKETS, "ActualxUp", "UpActual", "TargetxUp", "Up"]

    goals = _goal_map("AREA_LOSS_DAILY_GOALS", plant)
    no_good = [proc for proc, g in df.groupby("Process")["Good"].sum().items() if g <= 0]
    if no_good:
        log.info("Area loss %s: no good-module counts, left out: %s", plant["label"], ", ".join(sorted(no_good)))
    df = df[~df["Process"].isin(no_good)]
    by_tool = df.groupby(["ShiftDate", "Process", "Equipment"])[cols].sum()
    by_area = df.groupby(["ShiftDate", "Process"])[cols].sum()

    out = {}
    for (day, proc), r in by_area.iterrows():
        goal = goals.get(proc) or plant["daily_goal"]
        gap = goal - r["Good"]
        tools = []
        for equipment, t in by_tool.xs((day, proc), level=[0, 1]).iterrows():
            ct = t["ActualxUp"] / t["UpActual"] if t["UpActual"] else None
            target = t["TargetxUp"] / t["Up"] if t["Up"] else None
            tools.append(dict(tool=equipment, good=float(t["Good"]), ct=ct, ct_target=target,
                              **{b: float(t[b]) for b in LOSS_BUCKETS}))
      
        causes = sorted(((t[b], t["tool"], b, t) for t in tools for b in ("Downtime", "CT", "Quality") if t[b] > 0),
                        key=lambda c: c[0], reverse=True)
        capacity = float(r["Good"] + sum(r[b] for b in LOSS_BUCKETS))
        out.setdefault(day, {})[proc] = dict(
            good=float(r["Good"]), goal=goal, gap=gap, short=gap > AREA_LOSS_THRESHOLD * goal,
            capacity=capacity, below_goal=max(goal - capacity, 0.0),
            **{b: float(r[b]) for b in LOSS_BUCKETS}, tools=tools,
            causes=[dict(tool=tool, bucket=b, modules=float(m), ct=t["ct"], ct_target=t["ct_target"])
                    for m, tool, b, t in causes])
    return out


def _cause_text(c: dict) -> str:
    text = f"{c['tool']} {c['bucket'].lower() if c['bucket'] != 'CT' else 'CT'} {c['modules']:,.0f}"
    if c["bucket"] == "CT" and c["ct"] and c["ct_target"]:
        text += f" ({c['ct']:.1f} s vs {c['ct_target']:.1f} s target)"
    return text


def area_loss_highlights(losses: dict | None, plant: str, limit: int = 10) -> list[str]:
    """The worst area-days short of goal, with the tool losses behind them, for Copilot."""
    if not losses:
        return []
    short = sorted(((d, proc, a) for d, procs in losses.items() for proc, a in procs.items() if a["short"]),
                   key=lambda x: x[2]["gap"], reverse=True)
    if not short:
        return [f"Area loss at {plant}: every process made its daily goal "
                f"(within {AREA_LOSS_THRESHOLD:.0%}) every day."]
    lines = [f"Area loss at {plant}: {len(short)} process-days short of goal by more than "
             f"{AREA_LOSS_THRESHOLD:.0%}" + (f", worst {limit}:" if len(short) > limit else ":")]
    for day, proc, a in short[:limit]:
        own = ", ".join(_cause_text(c) for c in a["causes"][:AREA_LOSS_TOP_N])
        text = f"- {day:%a %m/%d} {proc}: {a['good']:,.0f} vs goal {a['goal']:,.0f} ({-a['gap']:+,.0f})"
        if a["below_goal"] > 0:
            text += (f"; can't make goal even at 100% MA and CT target: capacity {a['capacity']:,.0f}, "
                     f"{a['below_goal']:,.0f} short by design")
        text += f"; biggest tool losses (modules): {own}" if own else ""
        if a["Waiting"] > 0:
            text += f"; waiting on other areas {a['Waiting']:,.0f}"
        lines.append(text + ".")
    return lines


def constraint_keys(days: dict[date, dict] | None, plant: dict) -> set[str]:
    """Names a tool loss log Area can use for this week's constraint processes.

    The process name and its station names, e.g. Jbox and JUNCTION_BOX.
    """
    if not days:
        return set()
    names = constraint_tool_list(plant) or constraint_process_names(plant)[0]
    names = {n.split("-", 1)[-1]: proc for n, proc in names.items()}  # PGT14A-LAMINATOR -> LAMINATOR
    procs = {d["process"] for d in days.values()}
    keys = {_plant_key(p) for p in procs}
    keys |= {_plant_key(st) for st, proc in names.items() if proc in procs}
    keys |= {_plant_key(p.replace(" ", "_")) for p in procs}
    return keys


def constraint_highlights(days: dict[date, dict] | None, plant: str) -> list[str]:
    """Plain facts on the week's hourly constraint, for Copilot."""
    if not days:
        return []
    counts = pd.Series({}, dtype=float)
    for d in days.values():
        counts = counts.add(pd.Series(d["hours"], dtype=float), fill_value=0)
    counts = counts.sort_values(ascending=False)
    ct_kind = "actual CT" if CONSTRAINT_CT_SOURCE == "actual" else "CT target"
    lines = [f"Constraint at {plant}, hour by hour (lowest capacity at {ct_kind} among processes "
             f"not waiting), hours as constraint this week: "
             + ", ".join(f"{proc} {n:.0f} h" for proc, n in counts.items()) + "."]
    for day, d in sorted(days.items()):
        text = (f"- {day:%a %m/%d}: " + ", ".join(f"{p} {n} h" for p, n in d["hours"].items())
                + f"; line capacity {d['capacity']:,.0f} modules vs daily goal {d['goal']:,.0f}")
        text += " (short of goal)" if d["capacity"] < d["goal"] else ""
        if d["ct"] is not None and d["ct_target"]:
            text += (f"; {d['process']} ran at {d['ct']:.1f} s CT vs {d['ct_target']:.1f} s target, "
                     f"{d['tools']} tool{'s' if d['tools'] != 1 else ''}")
        if d["ma"] is not None:
            text += f", MA {d['ma']:.0%}"
        if d["idle"] is not None:
            text += f", blocked/starved {d['idle']:.0%} of uptime"
        if not d["confirmed"]:
            text += " (not confirmed: most hours every process was often waiting)"
        if d["margin"] is not None:
            text += f"; the next-lowest process had {d['margin']:,.0f} more over the day"
        if d["skipped"]:
            text += "; lower but waiting, passed over: " + ", ".join(
                f"{p} {n} h" for p, n in d["skipped"].items())
        lines.append(text + ".")
    return lines


def _constraint_tools(plant: dict) -> set[str]:
    """TOOL_LOSS_CONSTRAINT_TOOLS(_<prefix>): tool names separated by |, any case."""
    raw = _plant_setting("TOOL_LOSS_CONSTRAINT_TOOLS", plant)
    return {t.strip().lower() for t in raw.split("|") if t.strip()}


def tool_loss_metrics(df: pd.DataFrame | None, constraint_areas: dict[str, set[str]] | None = None) -> dict | None:
    """Per plant, only the tools worth reporting: big hitters and constraint tools.

    Big hitters are the plant's TOOL_LOSS_TOP_N tools with the most loss this
    week, ranked by the first metric (modules lost, else downtime).
    Constraint tools (TOOL_LOSS_CONSTRAINT_TOOLS, plus tools whose log Area
    matches a process that was the daily constraint this week, from
    `constraint_areas`) are added after them when they had any loss this
    week or last. A reported tool is on the watch list
    when its ranking metric rose TOOL_LOSS_WATCH_INCREASE (0.5 = 50%) or more.
    """
    if df is None:
        return None
    metrics = [m for m in TOOL_LOSS_METRICS if m in df.columns]
    rank = metrics[0]
    prev_window = report_window(days=2 * len(WINDOW))[:len(WINDOW)]
    week = df[df["ShiftDate"].isin(WINDOW)].groupby(["Plant", "Tool"])[metrics].sum()
    prev = df[df["ShiftDate"].isin(prev_window)].groupby(["Plant", "Tool"])[metrics].sum()

    def values(frame, key):
        return {x: float(frame.loc[key, x]) if key in frame.index else 0.0 for x in metrics}

    area_of = {}
    if "Area" in df.columns:
        named = df[df["Area"].ne("")]
        area_of = named.groupby(["Plant", "Tool"])["Area"].last().to_dict()

    plants = {}
    for p in PLANTS:
        label, constraints = p["label"], _constraint_tools(p)
        auto = (constraint_areas or {}).get(label, set())
        names = sorted({t for pl, t in week.index.union(prev.index) if pl == label})
        tools = []
        for name in names:
            cur, last = values(week, (label, name)), values(prev, (label, name))
            constraint = name.lower() in constraints or _plant_key(area_of.get((label, name), "")) in auto - {""}
            tools.append(dict(tool=name, cur=cur, prev=last, constraint=constraint))
        tools.sort(key=lambda t: t["cur"][rank], reverse=True)
        hitters = [t for t in tools if t["cur"][rank] > 0][:TOOL_LOSS_TOP_N]
        for t in hitters:
            t["big_hitter"] = True
        reported = hitters + [t for t in tools if t["constraint"] and t not in hitters
                              and (any(t["cur"].values()) or any(t["prev"].values()))]
        for t in reported:
            t["kind"] = ("Both" if t["constraint"] and t.get("big_hitter")
                         else "Constraint" if t["constraint"] else "Big hitter")
            t["watch"] = bool(t["prev"][rank]) and t["cur"][rank] >= t["prev"][rank] * (1 + TOOL_LOSS_WATCH_INCREASE)
        plants[label] = reported
        missing = constraints - {t["tool"].lower() for t in tools}
        if missing:
            log.info("Tool loss %s: constraint tools with no rows in the log: %s", label, ", ".join(sorted(missing)))
    return dict(metrics=metrics, plants=plants, prev_window=prev_window)


def tool_loss_label(metric: str) -> str:
    return "Modules lost" if metric == "ModulesLost" else f"Downtime ({TOOL_LOSS_DOWNTIME_UNIT})"


def _tool_loss_amount(metric: str, v: float) -> str:
    return f"{v:,.0f} modules lost" if metric == "ModulesLost" else f"{v:,.0f} {TOOL_LOSS_DOWNTIME_UNIT} of downtime"


def _vs_last_week(cur: float, prev: float) -> str:
    if not prev:
        return "none last week" if cur else "same as last week"
    return f"{'up' if cur >= prev else 'down'} {abs(cur / prev - 1):.0%} from {prev:,.0f} last week"


def _kind_text(t: dict) -> str:
    return {"Both": "big hitter, constraint tool", "Constraint": "constraint tool"}.get(t["kind"], "big hitter")


def tool_loss_highlights(m: dict, plant: str | None = None) -> list[str]:
    """Plain facts on the reported tools, one per line, for Copilot to write up.

    With `plant`, only that plant.
    """
    rank = m["metrics"][0]
    what = "modules lost" if rank == "ModulesLost" else "downtime"
    lines = [f"Tool loss big hitters and constraint tools, {week_label()}."]
    for name, tools in m["plants"].items():
        if plant and name != plant:
            continue
        lines.append(f"{name}:")
        if not tools:
            lines.append("- No big hitter or constraint tool loss logged this week.")
        for t in tools:
            lines.append(f"- {t['tool']} ({_kind_text(t)}): " + "; ".join(
                f"{_tool_loss_amount(x, t['cur'][x])} ({_vs_last_week(t['cur'][x], t['prev'][x])})"
                for x in m["metrics"]) + ".")
        for t in tools:
            if t["watch"]:
                lines.append(f"- Watch: {t['tool']} {what} up {t['cur'][rank] / t['prev'][rank] - 1:.0%} "
                             f"({t['prev'][rank]:,.0f} to {t['cur'][rank]:,.0f}).")
    return lines


def save_tool_loss_summary(m: dict | None, constraints: dict[str, dict] | None = None,
                           losses: dict[str, dict] | None = None) -> None:
    """Write the week's reported tools, constraints and area losses to TOOL_LOSS_SUMMARY_FILE.

    Save it in a OneDrive-synced SharePoint folder and point a Copilot agent
    at it. Overwritten every week, so it always holds the latest week.
    """
    constraints = {k: v for k, v in (constraints or {}).items() if v}
    losses = {k: v for k, v in (losses or {}).items() if v}
    if not TOOL_LOSS_SUMMARY_FILE or (m is None and not constraints and not losses):
        return
    week = week_label()
    sheets = {}
    if m is not None:
        sheets["Tools"] = pd.DataFrame([
        {"Week": week, "Plant": plant, "Tool": t["tool"], "Type": t["kind"],
         **{k: v for x in m["metrics"] for k, v in (
             (tool_loss_label(x), t["cur"][x]), (f"{tool_loss_label(x)} last week", t["prev"][x]))},
         "Watch": "Yes" if t["watch"] else ""}
        for plant, reported in m["plants"].items() for t in reported])
    if constraints:
        sheets["Daily Constraint"] = pd.DataFrame([
            {"Week": week, "Plant": plant, "Date": day, "Constraint (most hours)": d["process"],
             "Hours as constraint": ", ".join(f"{p} {n}" for p, n in d["hours"].items()),
             "Line capacity": round(d["capacity"]), "Daily goal": d["goal"],
             "Short of goal": "Yes" if d["capacity"] < d["goal"] else "",
             "Margin to next lowest": round(d["margin"]) if d["margin"] is not None else None,
             "Constraint CT (s)": d["ct"], "CT target (s)": d["ct_target"], "Tools": d["tools"],
             "MA": d["ma"], "OE": d["oe"], "Blocked + starved % of uptime": d["idle"],
             "Confirmed": "Yes" if d["confirmed"] else "No",
             "Lower but waiting (hours)": ", ".join(f"{p} {n}" for p, n in d["skipped"].items())}
            for plant, days in constraints.items() for day, d in sorted(days.items())])
        sheets["Hourly Constraint"] = pd.DataFrame([
            {"Plant": plant, "Date": day, "Hour": h["hour"], "Constraint": h["process"],
             "Capacity": round(h["capacity"]), "Next lowest": h["next"],
             "Next lowest capacity": round(h["next_capacity"]) if h["next"] else None,
             "Confirmed": "Yes" if h["confirmed"] else "No", "Passed over (waiting)": ", ".join(h["skipped"])}
            for plant, days in constraints.items() for day, d in sorted(days.items()) for h in d["hourly"]])
    if losses:
        sheets["Area Loss"] = pd.DataFrame([
            {"Week": week, "Plant": plant, "Date": day, "Process": proc, "Good modules": round(a["good"]),
             "Daily goal": a["goal"], "vs goal": round(-a["gap"]), "Short": "Yes" if a["short"] else "",
             "Capacity at CT target, 100% MA": round(a["capacity"]), "Short by design": round(a["below_goal"]),
             **{f"{b} loss": round(a[b]) for b in LOSS_BUCKETS},
             "Biggest tool losses": "; ".join(_cause_text(c) for c in a["causes"][:AREA_LOSS_TOP_N])}
            for plant, days in losses.items() for day, procs in sorted(days.items()) for proc, a in procs.items()])
        sheets["Area Loss by Tool"] = pd.DataFrame([
            {"Plant": plant, "Date": day, "Process": proc, "Tool": t["tool"], "Good modules": round(t["good"]),
             **{f"{b} loss": round(t[b]) for b in LOSS_BUCKETS},
             "Actual CT (s)": round(t["ct"], 1) if t["ct"] else None,
             "CT target (s)": round(t["ct_target"], 1) if t["ct_target"] else None}
            for plant, days in losses.items() for day, procs in sorted(days.items())
            for proc, a in procs.items() if a["short"] for t in a["tools"]])
    lines = tool_loss_highlights(m) if m is not None else []
    for plant, days in constraints.items():
        lines += constraint_highlights(days, plant)
    for plant, days in losses.items():
        lines += area_loss_highlights(days, plant, limit=50)
    sheets["Highlights"] = pd.DataFrame({"Highlight": lines})
    path = Path(TOOL_LOSS_SUMMARY_FILE)
    path = path if path.is_absolute() else SCRIPT_DIR / path
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with pd.ExcelWriter(path, engine="openpyxl") as xl:
            for name, frame in sheets.items():
                frame.to_excel(xl, sheet_name=name, index=False)
        log.info("Saved %s", path)
    except Exception as exc: 
        log.error("Could not save the tool loss summary %s: %s", path, exc)

# ---------------------------------------------------------------------------
# Formatting
# ---------------------------------------------------------------------------

def _num(v) -> str:
    return NOT_CONFIGURED if v is None else f"{v:,.0f}"


def _pct(v, places: int = 1) -> str:
    return NOT_CONFIGURED if v is None else f"{v:.{places}%}"


def goal_color(value, goal, higher_is_better: bool = True) -> RGBColor | None:
    """Green when the goal is met, yellow within GOAL_WARN_BAND of it, red past that."""
    if value is None or goal is None:
        return None
    if higher_is_better:
        if value >= goal:
            return COLOR_GOOD
        return COLOR_WARN if value >= goal * (1 - GOAL_WARN_BAND) else COLOR_BAD
    if value <= goal:
        return COLOR_GOOD
    return COLOR_WARN if goal and value <= goal * (1 + GOAL_WARN_BAND) else COLOR_BAD


def _goal(v, kind: str) -> str:
    if v is None:
        return NO_GOAL
    return f"{v:.2%}" if kind == "pct" else f"{v:,.0f}"


def week_label() -> str:
    start, end = WINDOW[0], WINDOW[-1]
    ww = start.isocalendar().week
    return f"WW{ww:02d} · {start:%b %d} – {end:%b %d, %Y}"



COLOR_NEGATIVE = RGBColor(0xC0, 0x00, 0x00)
COLOR_GOOD     = RGBColor(0x47, 0x8D, 0x34)   # the green, yellow and red the team uses
COLOR_WARN     = RGBColor(0xFF, 0xC0, 0x00)
COLOR_BAD      = RGBColor(0xFF, 0x00, 0x00)
COLOR_SUBTLE   = RGBColor(0x70, 0x70, 0x70)
PLANT_COLORS   = [RGBColor(0x2E, 0x9B, 0xF0), RGBColor(0x1F, 0x2A, 0x80), RGBColor(0xF0, 0x70, 0x30),
                  RGBColor(0x5B, 0xB5, 0x6E), RGBColor(0x8E, 0x6C, 0xC4)]


def _title_of(slide) -> str:
    title = slide.shapes.title
    return title.text_frame.text.strip() if title is not None and title.has_text_frame else ""


def _find_slide(prs, prefix: str, needs_table: bool = False):
    for slide in prs.slides:
        if _title_of(slide).lower().startswith(prefix.lower()):
            if not needs_table or any(s.has_table for s in slide.shapes):
                return slide
    return None


def _drop_slide(prs, slide) -> None:
    sld_ids = prs.slides._sldIdLst
    for sld_id in list(sld_ids):
        if prs.part.related_part(sld_id.rId) is slide.part:
            prs.part.drop_rel(sld_id.rId)
            sld_ids.remove(sld_id)
            return


def _add_slide(prs, layout):
    new = prs.slides.add_slide(layout)
   
    taken = {str(part.partname) for part in prs.part.package.iter_parts() if part is not new.part}
    n = len(prs.slides)
    while f"/ppt/slides/slide{n}.xml" in taken:
        n += 1
    new.part.partname = PackURI(f"/ppt/slides/slide{n}.xml")
    return new


def _duplicate_slide(prs, src):
    """Copy a slide made of placeholders, text and tables (no pictures/charts)."""
    new = _add_slide(prs, src.slide_layout)
    for shape in list(new.shapes):
        shape._element.getparent().remove(shape._element)
    for el in src.shapes._spTree.iterchildren():
        if el.tag.endswith(("}nvGrpSpPr", "}grpSpPr")):
            continue
        new.shapes._spTree.append(copy.deepcopy(el))
    return new


def _move_after(prs, slide, anchor) -> None:
    """Move `slide` so it sits right after `anchor` in the deck."""
    sld_ids = prs.slides._sldIdLst
    by_part = {prs.part.related_part(s.rId): s for s in sld_ids}
    el, anchor_el = by_part[slide.part], by_part[anchor.part]
    sld_ids.remove(el)
    anchor_el.addnext(el)


def _set_text(shape, text: str, color: RGBColor | None = None) -> None:
    """Replace a shape's or cell's text, keeping the first run's formatting."""
    tf = shape.text_frame
    paras = tf.paragraphs
    for p in paras[1:]:
        p._p.getparent().remove(p._p)
    p = paras[0]
    runs = p.runs
    if runs:
        runs[0].text = text
        for r in runs[1:]:
            r._r.getparent().remove(r._r)
        run = runs[0]
    else:
        run = p.add_run()
        run.text = text
        end = p._p.find("{http://schemas.openxmlformats.org/drawingml/2006/main}endParaRPr")
        if end is not None: 
            rpr = copy.deepcopy(end)
            rpr.tag = "{http://schemas.openxmlformats.org/drawingml/2006/main}rPr"
            run._r.insert(0, rpr)
    if color is not None:
        run.font.color.rgb = color
    else: 
        rpr = run._r.find("{http://schemas.openxmlformats.org/drawingml/2006/main}rPr")
        if rpr is not None:
            for fill in rpr.findall("{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill"):
                rpr.remove(fill)


def _body_placeholder(slide):
    for shape in slide.placeholders:
        if shape.placeholder_format.type == PP_PLACEHOLDER.BODY:
            return shape
    return None


def _set_source(slide, text: str, warn: bool = False) -> None:
    ph = _body_placeholder(slide)
    if ph is None:
        ph = slide.shapes.add_textbox(Inches(0.5), Inches(5.1), Inches(9.1), Inches(0.3))
        ph.text_frame.word_wrap = True
        ph.text_frame.text = " "
        ph.text_frame.paragraphs[0].runs[0].font.size = Pt(9)
    elif ph.top < Inches(4.5):  # some template slides hold notes mid-slide, under the table
        ph.top, ph.height = Inches(5.1), Inches(0.3)
    _set_text(ph, text, COLOR_NEGATIVE if warn else None)


def _first_table(slide):
    return next((s.table for s in slide.shapes if s.has_table), None)


def _fill_table(table, rows: list[list[str]], colors: list[list] | None = None) -> None:
    """Write rows under the header row; spare template rows are blanked.

    `colors` matches `rows` cell for cell (None = the table's own color).
    """
    body = len(table.rows) - 1
    if len(rows) > body:
        log.warning("Template table has %d rows, %d needed — extra rows dropped.", body, len(rows))
    for ri in range(body):
        values = rows[ri] if ri < len(rows) else [""] * len(table.columns)
        for ci in range(len(table.columns)):
            text = values[ci] if ci < len(values) else ""
            color = None
            if colors and ri < len(colors) and ci < len(colors[ri]):
                color = colors[ri][ci]
            if text in (NOT_CONFIGURED, NO_GOAL):
                color = COLOR_SUBTLE
            _set_text(table.cell(ri + 1, ci), text, color)

def _insert_column(table, before: int, header: str) -> None:
    """Insert a column before `before`, styled like that column, with `header` on top."""
    grid = table._tbl.tblGrid
    grid.insert(before, copy.deepcopy(grid[before]))
    for tr in table._tbl.tr_lst:
        tcs = tr.tc_lst
        tcs[before].addprevious(copy.deepcopy(tcs[before]))
    _set_text(table.cell(0, before), header)


def _set_column_widths(table, inches: list[float]) -> None:
    for col, w in zip(table.columns, inches):
        col.width = Inches(w)
    for row in table.rows:
        row.height = Inches(0.4)

# ---------------------------------------------------------------------------
# Slide builders
# ---------------------------------------------------------------------------

def build_title_slide(slide) -> None:
    title = slide.shapes.title
    if title is not None:
        _set_text(title, "Weekly Operations Report")
    for ph in slide.placeholders:
        if ph.placeholder_format.type == PP_PLACEHOLDER.BODY and ph.height < Inches(2):
            _set_text(ph, week_label())
            break



def build_throughput_slide(slide, fg_by_plant: dict, areas_by_plant: dict) -> None:
    """Chart 1: FG per shift day, stacked by plant. Chart 2: area daily average by plant."""
    for shape in list(slide.shapes):  # sample picture or last week's charts
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE or getattr(shape, "has_chart", False):
            shape._element.getparent().remove(shape._element)

    fg = CategoryChartData()
    fg.categories = [f"{d:%a %m/%d}" for d in WINDOW]
    for p in PLANTS:
        days = fg_by_plant.get(p["label"])
        fg.add_series(p["label"], [days[d] if days else None for d in WINDOW])
    target = fg_daily_target()
    if target:
        fg.add_series(f"Target {target:,.0f}", [target] * len(WINDOW))
    chart = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, Inches(0.3), Inches(1.0),
                                   Inches(5.6), Inches(4.0), fg).chart
    if target:
        _last_series_as_line(chart, COLOR_BAD)
    _style_chart(chart, "FG per day by plant", labels_inside=True)
    if FG_CHART_MAX.strip():
        try:
            chart.value_axis.maximum_scale = float(FG_CHART_MAX)
        except ValueError:
            log.warning("FG_CHART_MAX=%r is not a number — PowerPoint picks the axis.", FG_CHART_MAX)

    area_names = [a for a in area_stations(PLANTS[0]) if a != "Bussing"] if PLANTS else []
    ar = CategoryChartData()
    ar.categories = area_names
    for p in PLANTS:
        metrics = areas_by_plant.get(p["label"]) or {}
        ar.add_series(p["label"], [(metrics.get(a) or {}).get("avg") for a in area_names])
    chart = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(6.0), Inches(1.0),
                                   Inches(3.8), Inches(4.0), ar).chart
    _style_chart(chart, "LAM, Jbox and Frame, avg per day")

    missing = [p["label"] for p in PLANTS if not fg_by_plant.get(p["label"])]
    note = f"Week of {WINDOW[0]:%a %b %d} – {WINDOW[-1]:%a %b %d}, midnight to midnight."
    if missing:
        note += f" No data for {', '.join(missing)} (SQL server not configured or query failed)."
    _set_source(slide, note, warn=bool(missing))


def fg_daily_target() -> float | None:
    """FG_DAILY_TARGET, or the sum of the plants' daily FG goals."""
    if FG_DAILY_TARGET.strip():
        try:
            return float(FG_DAILY_TARGET) or None
        except ValueError:
            log.warning("FG_DAILY_TARGET=%r is not a number — no target line.", FG_DAILY_TARGET)
            return None
    return sum(p["daily_goal"] for p in PLANTS) or None


def _last_series_as_line(chart, color: RGBColor) -> None:
    """Move the chart's last bar series into a line plot on the same axes."""
    bar = chart.plots[0]._element
    ser = bar.findall(qn("c:ser"))[-1]
    bar.remove(ser)
    for tag in ("c:spPr", "c:invertIfNegative"):
        for el in ser.findall(qn(tag)):
            ser.remove(el)
    line = bar.makeelement(qn("c:lineChart"), {})
    for tag, val in (("c:grouping", "standard"), ("c:varyColors", "0")):
        line.append(line.makeelement(qn(tag), {"val": val}))
    line.append(ser)
    sppr = ser.makeelement(qn("c:spPr"), {})
    ln = sppr.makeelement(qn("a:ln"), {"w": "22225"})
    fill = ln.makeelement(qn("a:solidFill"), {})
    fill.append(fill.makeelement(qn("a:srgbClr"), {"val": str(color)}))
    ln.append(fill)
    ln.append(ln.makeelement(qn("a:prstDash"), {"val": "dash"}))
    sppr.append(ln)
    ser.find(qn("c:tx")).addnext(sppr)
    marker = ser.makeelement(qn("c:marker"), {})
    marker.append(marker.makeelement(qn("c:symbol"), {"val": "none"}))
    sppr.addnext(marker)
    ser.append(ser.makeelement(qn("c:smooth"), {"val": "0"}))
    line.append(line.makeelement(qn("c:marker"), {"val": "1"}))
    for ax in bar.findall(qn("c:axId")):
        line.append(copy.deepcopy(ax))
    bar.addnext(line)


def _style_chart(chart, title: str, labels_inside: bool = False) -> None:
    chart.has_title = True
    chart.chart_title.text_frame.text = title
    chart.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(11)
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    chart.legend.font.size = Pt(9)
    chart.category_axis.tick_labels.font.size = Pt(8)
    chart.value_axis.tick_labels.font.size = Pt(8)
    chart.value_axis.has_major_gridlines = False
    chart.value_axis.minimum_scale = 0
    for i, series in enumerate(chart.plots[0].series):
        series.format.fill.solid()
        series.format.fill.fore_color.rgb = PLANT_COLORS[i % len(PLANT_COLORS)]
    plot = chart.plots[0]
    plot.has_data_labels = True
    labels = plot.data_labels
    labels.number_format = "#,##0"
    labels.number_format_is_linked = False
    labels.font.size = Pt(7)
    if labels_inside:  # stacked segments: white reads on every plant color
        labels.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        labels.font.bold = True


def build_critical_area_slide(slide, plant: dict, metrics: dict) -> None:
    _set_text(slide.shapes.title, f"Throughput Critical Area: {plant['label']}")
    ma_goals = _goal_map("AREA_MA_GOALS", plant)
    weekly_goals = _goal_map("AREA_WEEKLY_GOALS", plant)
    rows, colors = [], []
    for area, m in metrics.items():
        if m is None:
            rows.append([area] + [NOT_CONFIGURED] * 5)
            colors.append([None] * 6)
            continue
        weekly = weekly_goals.get(area)
        daily = weekly / len(WINDOW) if weekly else None
        basis = f"{_num(m['total'])} / {_num(weekly)}" if weekly else f"{_num(m['total'])} / {NO_GOAL}"
        rows.append([area, _pct(m["ma"]), _num(m["avg"]), _num(m["median"]), basis,
                     _pct(m["utilization"]) if m["utilization"] is not None else NO_GOAL])
        colors.append([None, goal_color(m["ma"], ma_goals.get(area)), goal_color(m["avg"], daily),
                       goal_color(m["median"], daily), None, goal_color(m["utilization"], 1.0)])
    table = _first_table(slide)
    if table is not None:
        if len(table.columns) < 6:  
            _insert_column(table, 4, "Week total / Goal")
        _fill_table(table, rows, colors)
        _set_column_widths(table, [1.4, 1.0, 1.2, 1.4, 2.2, 1.5])
    _set_source(slide, "MA = uptime / scheduled minutes. Avg and Median P50 = modules per shift day. "
                       "Utilization % = week total / weekly goal x 100. "
                       "Green meets goal, yellow is close, red misses. "
                       f"{NO_GOAL} = no weekly goal in config.env, {NOT_CONFIGURED} = no data source.")


def build_crew_slide(slide, plant: dict, data: dict) -> None:
    _set_text(slide.shapes.title, f"Crew Performance: {plant['label']}")
    goals = _goal_map("CREW_GOALS", plant)

    colors = []

    def row(label, key, values, kind, higher_is_better=True):
        goal = goals.get(key)
        cells = [label, _goal(goal, kind)]
        colors.append([None, COLOR_SUBTLE] + [
            goal_color(None if values is None else values.get(crew), goal, higher_is_better)
            for crew in CREWS])
        for crew in CREWS:
            v = None if values is None else values.get(crew)
            if values is None:
                cells.append(NOT_CONFIGURED)
            elif v is None:
                cells.append("–")
            else:
                cells.append(f"{v:.2%}" if kind == "pct" else f"{v:,.0f}")
        return cells

    per_shift = CREW_THROUGHPUT_MODE == "per_shift"
    rows = [
        row("Frame Throughput per shift" if per_shift else "Frame Throughput",
            "Throughput", data["throughput"], "num"),
        row("Rework rerun per shift", "ReworkRerun", data["rework_rerun"], "pct", False),
        row("Frame scrap", "FrameScrap", data["frame_scrap"], "pct", False),
        row("Catchall", "Catchall", data["catchall"], "num", False),
        row("FG", "FG", data["fg"], "num"),
    ]
    table = _first_table(slide)
    if table is not None:
        header = ["Metric", "Goal"] + CREWS
        for ci in range(min(len(header), len(table.columns))):
            _set_text(table.cell(0, ci), header[ci])
        _fill_table(table, rows, colors)
        crew_w = (9.2 - 2.2 - 1.0) / max(len(CREWS), 1)
        _set_column_widths(table, [2.2, 1.0] + [crew_w] * len(CREWS))
    _set_source(slide, f"Throughput {'per shift ' if per_shift else ''}at {CREW_THROUGHPUT_AREA}. "
                       "Rework rerun = Hipot reruns / modules tested. Green meets goal, yellow is close, red misses. "
                       f"{len(CREW_WINDOW)} shift days ({SHIFT_START_HOUR}:00 start), {CREW_WINDOW[0]:%a %b %d} – {CREW_WINDOW[-1]:%a %b %d}. "
                       f"{NOT_CONFIGURED} = no data source yet, {NO_GOAL} = no goal or no data for that crew.")


def _add_table(slide, rows: list[list[str]], top: float, widths: list[float],
               colors: list[list] | None = None, size: int = 10, left: float = 0.4,
               row_height: float = 0.28, name: str = ""):
    """Add a table (first row = header) at `top` inches, columns `widths` inches wide."""
    shape = slide.shapes.add_table(len(rows), len(widths), Inches(left), Inches(top),
                                   Inches(sum(widths)), Inches(row_height * len(rows)))
    if name:
        shape.name = name
    table = shape.table
    for ci, w in enumerate(widths):
        table.columns[ci].width = Inches(w)
    for ri, values in enumerate(rows):
        table.rows[ri].height = Inches(row_height)
        for ci, text in enumerate(values):
            cell = table.cell(ri, ci)
            cell.margin_top = cell.margin_bottom = Inches(0.02)
            cell.margin_left = cell.margin_right = Inches(0.06)
            cell.text = text
            run = cell.text_frame.paragraphs[0].runs[0] if text else None
            if run is None:
                continue
            run.font.size = Pt(size)
            color = colors[ri][ci] if colors and ri < len(colors) and ci < len(colors[ri]) else None
            if text in (NOT_CONFIGURED, NO_GOAL):
                color = COLOR_SUBTLE
            if color is not None:
                run.font.color.rgb = color
    return shape


def _change_cell(cur: float, prev: float) -> tuple[str, RGBColor | None]:
    """'+27%' in red when up on last week, green when down."""
    if not prev:
        return ("new", COLOR_BAD) if cur else (NO_GOAL, None)
    return f"{cur / prev - 1:+.0%}", COLOR_BAD if cur > prev else COLOR_GOOD


PLANT_TOOL_LOSS_TABLE = "Tool Loss by Tool"
CONSTRAINT_TABLE      = "Daily Constraint"
WHAT_HAPPENED_BOX     = "What happened"
WHAT_HAPPENED_WIDTH   = 3.0  


def _tables_bottom(slide, skip: tuple[str, ...]) -> float | None:
    """Bottom edge (inches) of the lowest table on the slide, ignoring `skip` names."""
    bottoms = [(sh.top + sum(r.height for r in sh.table.rows)) / 914400
               for sh in slide.shapes if sh.has_table and sh.name not in skip]
    return max(bottoms) if bottoms else None


def build_constraint_row(slide, plant: dict, days: dict[date, dict] | None) -> None:
    """Under the Critical Area table: each shift day's constraint and the line's capacity.

    The process is the one that was the constraint for the most hours, with
    its hours. Capacity is each hour's constraint capacity added up: green
    when it covers the plant's daily goal, red when it can't; (+N) is how
    much more the next-lowest process had, over the day. '?' marks a day
    when most hours every process was often blocked or starved.
    """
    for shape in list(slide.shapes):  
        if shape.name == CONSTRAINT_TABLE:
            shape._element.getparent().remove(shape._element)
    if days is None and not CONSTRAINT_QUERY_FILE:
        return
    area = next((sh for sh in slide.shapes if sh.has_table), None)
    left = area.left / 914400 if area is not None else 0.6
    width = area.width / 914400 if area is not None else 8.7
    top = (_tables_bottom(slide, (PLANT_TOOL_LOSS_TABLE,)) or 3.0) + 0.12
    days = days or {}
    head, procs, caps = ["Constraint"], ["Process"], ["Capacity"]
    colors = [[], [None], [None]]
    for d in WINDOW:
        c = days.get(d)
        head.append(f"{d:%a %m/%d}")
        if c is None:
            procs.append(NOT_CONFIGURED)
            caps.append(NOT_CONFIGURED)
            colors[1].append(None)
            colors[2].append(None)
            continue
        procs.append(f"{c['process']} {c['hours'][c['process']]}h" + ("" if c["confirmed"] else " ?"))
        margin = f" (+{c['margin']:,.0f})" if c["margin"] is not None else ""
        caps.append(f"{c['capacity']:,.0f}{margin}")
        colors[1].append(None)
        colors[2].append(COLOR_BAD if c["capacity"] < c["goal"] else COLOR_GOOD)
    label_w = 0.95
    _add_table(slide, [head, procs, caps], top, [label_w] + [(width - label_w) / len(WINDOW)] * len(WINDOW),
               colors, size=8, left=left, row_height=0.22, name=CONSTRAINT_TABLE)


def what_happened_draft(m: dict | None, plant: str, days: dict[date, dict] | None = None) -> list[str]:
    """Up to three one-line bullets on the plant's constraint and tools, for Copilot to rewrite.

    Plain enough to present as is when nobody rewrites it. The table
    beside it has the detail, so each bullet names one thing.
    """
    lines = []
    if days:
        counts = pd.Series([d["process"] for d in days.values()]).value_counts()
        short = sum(d["capacity"] < d["goal"] for d in days.values())
        lines.append(f"Constraint (most hours): {counts.index[0]}, {counts.iloc[0]} of {len(days)} days; "
                     + (f"short of goal {short} days." if short else "met goal every day."))
    if m is None:
        return lines
    rank, tools = m["metrics"][0], m["plants"][plant]
    unit = "modules" if rank == "ModulesLost" else TOOL_LOSS_DOWNTIME_UNIT
    if not tools:
        return lines + ["No big hitter or constraint tool loss."]
    for t in [t for t in tools if t["watch"]][:2]:
        lines.append(f"{t['tool']} up {t['cur'][rank] / t['prev'][rank] - 1:.0%} "
                     f"({t['prev'][rank]:,.0f} to {t['cur'][rank]:,.0f} {unit}).")
    if not tools[0]["watch"]:
        lines.insert(1 if days else 0, f"Biggest hitter: {tools[0]['tool']}, "
                                       f"{tools[0]['cur'][rank]:,.0f} {unit}.")
    return lines[:3]


def _add_what_happened(slide, lines: list[str], left: float, top: float, width: float, height: float) -> None:
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    box.name = WHAT_HAPPENED_BOX
    tf = box.text_frame
    tf.word_wrap = True
    head = tf.paragraphs[0].add_run()
    head.text = "What happened"
    head.font.size, head.font.bold = Pt(11), True
    for line in lines:
        para = tf.add_paragraph()
        para.space_before = Pt(4)
        run = para.add_run()
        run.text = f"• {line}"
        run.font.size = Pt(9)


def build_plant_tool_loss(slide, plant: dict, m: dict | None, days: dict[date, dict] | None = None) -> None:
    """Under the Critical Area table (and constraint row): big hitter and constraint tools.

    Rows that don't fit above the source line are left off. Beside the table,
    a "What happened" box holds a draft from the numbers for Copilot to
    rewrite; the plant's full facts go in the speaker notes.
    """
    for shape in list(slide.shapes): 
        if shape.name in (PLANT_TOOL_LOSS_TABLE, WHAT_HAPPENED_BOX):
            shape._element.getparent().remove(shape._element)
    area = next((s for s in slide.shapes if s.has_table), None)
    left = area.left / 914400 if area is not None else 0.6
    width = area.width / 914400 if area is not None else 8.7
    top = (_tables_bottom(slide, (PLANT_TOOL_LOSS_TABLE,)) or 3.0) + 0.15
    row_h = 0.22
    fit = max(int((5.0 - top) / row_h) - 1, 0)  
    if m is None:
        if TOOL_LOSS_FILE: 
            box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(0.3))
            box.name = PLANT_TOOL_LOSS_TABLE
            _set_text(box, f"Tool loss {NOT_CONFIGURED}: the tool loss log couldn't be read "
                           "(see the run log).", COLOR_SUBTLE)
            box.text_frame.paragraphs[0].runs[0].font.size = Pt(10)
        return
    if top > 4.5:
        log.warning("No room for the tool loss table on %s's Critical Area slide.", plant["label"])
        return

    metrics, label = m["metrics"], plant["label"]
    tools = m["plants"][label]
    if len(tools) > fit:
        log.warning("Tool loss %s: %d tools to report, room for %d — the rest are in the speaker notes.",
                    label, len(tools), fit)
        # Short of rows, constraint tools stay and the smallest big hitters go.
        keep = [t for t in tools if t["constraint"]][:fit]
        keep += [t for t in tools if not t["constraint"]][:fit - len(keep)]
        tools = [t for t in tools if t in keep]
    short = {"ModulesLost": "Modules", "Downtime": f"Down {TOOL_LOSS_DOWNTIME_UNIT}"}
    rows = [["Tool", "Type"] + [h for x in metrics for h in (short[x], "vs LW")]]
    colors = [[]]
    for t in tools[:fit]:
        row, color = [t["tool"], t["kind"]], [None, COLOR_SUBTLE]
        for x in metrics:
            change, change_color = _change_cell(t["cur"][x], t["prev"][x])
            row += [f"{t['cur'][x]:,.0f}", change]
            color += [None, change_color]
        rows.append(row)
        colors.append(color)
    if not tools:
        rows.append(["No big hitter or constraint tool loss", ""] + [""] * (2 * len(metrics)))
    table_w = width - WHAT_HAPPENED_WIDTH - 0.15
    first, kind = min(1.45, table_w / 3), 0.8
    rest = (table_w - first - kind) / (2 * len(metrics))
    _add_table(slide, rows, top, [first, kind] + [rest] * (2 * len(metrics)), colors, size=9,
               left=left, row_height=row_h, name=PLANT_TOOL_LOSS_TABLE)
    _add_what_happened(slide, what_happened_draft(m, label, days), left + table_w + 0.15, top,
                       WHAT_HAPPENED_WIDTH, 5.0 - top)


def set_critical_notes(slide, plant: dict, m: dict | None, days: dict[date, dict] | None,
                       losses: dict | None = None) -> None:
    """The plant's constraint, area loss and tool loss facts, for Copilot, in the speaker notes."""
    lines = constraint_highlights(days, plant["label"]) + area_loss_highlights(losses, plant["label"])
    if m is not None:
        lines += tool_loss_highlights(m, plant["label"])
    if not lines:
        return
    slide.notes_slide.notes_text_frame.text = "\n".join(
        ["Copilot: rewrite the 'What happened' box from these facts, three short bullets, "
         "only these numbers.", ""] + lines)

# ---------------------------------------------------------------------------
# Deck assembly + save
# ---------------------------------------------------------------------------


DROP_TITLES = ("Yield", "Icons can be recolored", "Place icons", "Tool Loss")


def _open_template() -> Presentation:
    path = (SCRIPT_DIR / PPTX_TEMPLATE).resolve()
    if not path.exists():
        raise FileNotFoundError(f"PPTX_TEMPLATE not found: {path}")
    log.info("Using template: %s", path)
    return Presentation(str(path))


def build_deck() -> Presentation:
    prs = _open_template()

    critical = _find_slide(prs, "Throughput Critical", needs_table=True)
    crew = _find_slide(prs, "Crew Performance", needs_table=True)
    for slide in list(prs.slides):
        title = _title_of(slide)
      
        spare = ((title.startswith("Crew Performance") and slide is not crew)
                 or (title.startswith("Throughput Critical") and slide is not critical))
        if spare or title.startswith(DROP_TITLES):
            _drop_slide(prs, slide)

    build_title_slide(prs.slides[0])

    rework = load_rework_rerun()
    fg_by_plant, areas_by_plant, crew_by_plant, constraint_by_plant, loss_by_plant = {}, {}, {}, {}, {}
    for p in PLANTS:
        hours = fetch_station_hours(p)
        tool_hours = fetch_constraint_data(p)
        constraint_by_plant[p["label"]] = daily_constraints(tool_hours, p)
        loss_by_plant[p["label"]] = area_losses(tool_hours, p)
        fg_by_plant[p["label"]] = fetch_fg_daily(p)
        areas_by_plant[p["label"]] = area_metrics(hours, p)
        rr = fetch_rework_rerun(p)
        if rr is None and rework is not None:
            sub = rework[rework["Plant"] == p["label"]]
            rr = {r["Crew"]: r["Value"] for _, r in sub.iterrows() if pd.notna(r["Value"])}
        crew_by_plant[p["label"]] = dict(
            throughput=crew_throughput(hours),
            rework_rerun=rr,
            frame_scrap=fetch_frame_scrap(p),
            catchall=fetch_catchall(p),
            fg=fetch_crew_fg(p),
        )

    throughput = _find_slide(prs, "Throughput")
    if throughput is not None and throughput is not critical:
        build_throughput_slide(throughput, fg_by_plant, areas_by_plant)

    tool_loss = tool_loss_metrics(load_tool_loss(), {
        p["label"]: constraint_keys(constraint_by_plant[p["label"]], p) for p in PLANTS})

 
    for template_slide, builder, per_plant in ((critical, build_critical_area_slide, areas_by_plant),
                                               (crew, build_crew_slide, crew_by_plant)):
        if template_slide is None:
            log.warning("Template is missing a table slide for %s", builder.__name__)
            continue
        copies = [template_slide] + [_duplicate_slide(prs, template_slide) for _ in PLANTS[1:]]
        anchor = template_slide
        for slide in copies[1:]:
            _move_after(prs, slide, anchor)
            anchor = slide
        for slide, p in zip(copies, PLANTS):
            builder(slide, p, per_plant[p["label"]])
            if builder is build_critical_area_slide:
                days = constraint_by_plant[p["label"]]
                build_constraint_row(slide, p, days)
                build_plant_tool_loss(slide, p, tool_loss, days)
                set_critical_notes(slide, p, tool_loss, days, loss_by_plant[p["label"]])
    return prs, tool_loss, constraint_by_plant, loss_by_plant


def run() -> None:
    prs, tool_loss, constraints, losses = build_deck()
    save_tool_loss_summary(tool_loss, constraints, losses)
    out_dir = Path(PPTX_OUTPUT_DIR) if PPTX_OUTPUT_DIR else SCRIPT_DIR
    out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / f"Weekly_Ops_Report_{date.today():%Y-%m-%d}.pptx"
    prs.save(str(out))
    log.info("Saved %s (%d slides)", out, len(prs.slides))


if __name__ == "__main__":
    run()
