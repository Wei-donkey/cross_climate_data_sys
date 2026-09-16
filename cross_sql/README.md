# CROSS Database Schema (DDL)

Exported Oracle DDL (`CREATE TABLE ...` statements, via
`DBMS_METADATA.GET_DDL`) for every table under the four database accounts
CROSS uses. This documents the schema that [`cross_database`](../cross_database)
writes to and [`cross_desktop`](../cross_desktop) reads from — independent of
either codebase, and useful as a reference even without access to the live
database.

One subfolder per account (account/ini section names come from
`cross_database/scripts/Config_DB.ini.example`):

| Folder | Oracle account | Tables |
|---|---|---|
| `cross_climate/` | `climate` | 44 |
| `cross_hour/` | `weather` | 2 representative tables (see below) |
| `cross_meta/` | `cross_meta` | 10 |
| `cross_rain/` | `climate_rain` | 7 |

## `cross_hour/`

The `weather` account has well over a hundred tables, nearly all of them
per-year copies of the same two table designs (surface and AWST hourly
observations). Rather than exporting every year's identical structure, only
one representative table per design is included; `cross_hour/README.md`
lists the rest of the same-design tables by name.

## Regenerating

DDL was pulled with a small Python script using the `oracledb` driver
(thin mode — no Oracle Instant Client needed) and
`DBMS_METADATA.SET_TRANSFORM_PARAM` to suppress storage/tablespace/segment
noise from the output. Connect using each account's credentials from
`cross_database/scripts/Config_DB.ini`, then for each table:

```sql
SELECT DBMS_METADATA.GET_DDL('TABLE', :table_name, :owner) FROM dual;
```

## Note

The primary-key constraints in `cross_meta/` are named `PK_<table_name>` for
readability; this was added when writing these files and is documentation
only — it doesn't reflect a rename applied to the live database. All other
identifiers (table/column/constraint names) reflect the live schema
verbatim.
