# CROSS Climate Data Pipeline

Python scripts that pull/QC/aggregate surface and automatic-weather-station (AWST)
climate observation data for the CROSS database, sourced from the CROSS Oracle
databases and the MUSIC interface platform, and write the results back to Oracle.

These scripts run as scheduled jobs (Windows Task Scheduler) on the production
server; this repo is a working copy shared for reference/review.

## Setup

1. Install dependencies — either:
   - conda: `conda env create -f environment.yml && conda activate cross_database`, or
   - pip: `pip install -r requirements.txt`
2. Copy the two example config files (in `scripts/`) and fill in real
   credentials:
   ```
   cd scripts
   cp Config_DB.ini.example Config_DB.ini
   cp Config_MUSIC_GD.ini.example Config_MUSIC_GD.ini
   ```
3. `scripts/Config_Vars.ini` (column mappings / thresholds) is not sensitive
   and is committed as-is.

## Layout

All pipeline code and config live under `scripts/`; each top-level folder
there is one stage of the pipeline:

| Folder | Purpose |
|---|---|
| `music_station_info` | Pull/export station metadata from the MUSIC interface |
| `music_awst` | Pull hourly AWST (automatic weather station) obs from the MUSIC interface into Oracle |
| `music_surf` | Pull hourly/daily surface obs from the MUSIC interface into Oracle |
| `music_surf_awst_pre_xtrm` | Pull + QC precipitation extremes from the MUSIC interface |
| `stat_awst`, `stat_awst_xtrm` | Aggregate AWST data to day/dekad (ten)/month/quarter/year, incl. extremes |
| `stat_surf`, `stat_surf_xtrm`, `stat_surf_norm` | Aggregate national-station surface data (incl. extremes, climate normals) |
| `stat_surf_awst_ca` | Cold-air (cold surge) process statistics |
| `stat_surf_awst_rhre` | Regional heavy rainfall event (RHRE) statistics |
| `stat_surf_rsod` | Rainfall season onset date (RSOD) statistics |
| `common` | Shared helpers: `config.py` (ini loading), `db.py` (Oracle engine creation) |

`music_awst` and `music_surf` each also carry one incremental-update aggregation
script alongside the pull scripts (`stat_awst_cli_mul_day_update.py` /
`stat_surf_cli_mul_day_update.py`).

**Filename suffixes you'll see repeated across folders**: `_xtrm` = extremes,
`_norm` = climate normals, `_hk_mc` = Hong Kong/Macau-only variant, `_ext` =
wider region variant, `_station` = single-station rerun/backfill
copy, `_update` = incremental/backfill variant of the base script. These are
parameter variants of a base script rather than distinct logic.

## `common/` module

- `common/config.py` — `load_ini(filename)`: reads an ini file from
  `scripts/` (its own parent's parent — i.e. `common/`'s sibling level)
  regardless of the caller's working directory.
- `common/db.py` — `get_engine(section, ini_file='Config_DB.ini')`: builds the
  SQLAlchemy Oracle engine for a `Config_DB.ini` section.
- `common/__init__.py` marks the folder as an importable package; it lives
  inside `common/` (not `scripts/` itself) because that's what makes
  `from common.db import get_engine` resolve — each script adds `scripts/`
  (its own parent) to `sys.path` first, then imports `common` as a package
  from there. This is why `common/` and the `Config_*.ini` files must stay
  siblings of the pipeline folders inside `scripts/` — moving only one of
  them would break every script's imports and ini lookups.

Every script under this project now uses these two helpers instead of
repeating the ini-parsing/connection-string boilerplate inline.

## Known issues / follow-ups

- **Hardcoded date overrides**: many scripts compute `date_end =
  datetime.datetime.now()` and then immediately overwrite it with a hardcoded
  literal date on the next line. This is intentional — it's how specific
  historical periods get backfilled — not a bug to fix.
- **No logging/error handling**: scripts use `print()` for status and mostly
  have no try/except. This is a known, accepted tradeoff for this codebase.
- `sqlalchemy` is pinned below 2.0 because scripts call the 1.x-style
  `Engine.execute(sql)` directly, which was removed in SQLAlchemy 2.0. Not a
  concern while running in the current closed network with pinned dependency
  versions.
