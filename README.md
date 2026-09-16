# CROSS — Climatic Re-statistical-data Operation and Service System

CROSS is a climate-statistics platform used by meteorological departments
across Guangdong Province: a Windows desktop client for interactive
querying, mapping, and charting of climate/weather data, backed by an
Oracle database and a set of scheduled data pipeline jobs. This repository
packages the three pieces of that system together for portfolio/reference
purposes.

## Layout

| Folder | What it is |
|---|---|
| [`cross_desktop`](cross_desktop) | VB6 Windows desktop client. `src/` holds the VB6 project (`CROSS.vbp`) plus the runtime asset folders it depends on at run time (`Ini/`, `Template/`, `Preview/`, etc.). Lets users query, map (via Surfer), chart surface/AWST climate data, and run statistical aggregations. See [`cross_desktop/README.md`](cross_desktop/README.md), which includes a full screenshot walkthrough of every feature. |
| [`cross_database`](cross_database) | Python data pipeline. `scripts/` holds all pipeline code and config; the folder root just has `README.md`/dependency files. Pulls/QCs/aggregates surface and automatic-weather-station (AWST) observations from the MUSIC interface platform into Oracle, and runs as scheduled jobs on the production server. See [`cross_database/README.md`](cross_database/README.md) for the full pipeline layout. |
| [`cross_sql`](cross_sql) | Exported Oracle DDL (`CREATE TABLE` statements) for every table under the four database accounts the system uses, one subfolder per account: `cross_climate`, `cross_hour`, `cross_meta`, `cross_rain`. Documents the schema `cross_database` and `cross_desktop` both read/write, independent of either codebase. See [`cross_sql/README.md`](cross_sql/README.md). |

## How the pieces fit together

- `cross_desktop` and `cross_database` both connect to the same Oracle
  database, whose schema is captured in `cross_sql`.
- `cross_database` is the ingestion/aggregation layer: it populates the
  tables that `cross_desktop` later queries, maps, and charts for end users.
- Both `cross_desktop` and `cross_database` keep real credentials and
  internal server addresses out of version control (`.gitignore` + `*.example`
  template files alongside them, under `cross_desktop/src/` and
  `cross_database/scripts/` respectively) — copy the relevant `*.example`
  file, and fill in real values locally.

## License

[MIT](LICENSE)