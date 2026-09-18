# CROSS Desktop Client

The VB6 Windows desktop client for CROSS ("Climatic Re-statistical-data
Operation and Service System") — the interactive front end meteorological
staff use to query, map, and chart climate data from the Oracle database
that [`cross_database`](../cross_database) populates. (See the project
[root `README.md`](../README.md) for how this fits alongside
`cross_database` and `cross_sql`.)

## Layout

- `CROSS2.14En_Installer.exe` — the compiled installer for end users. It can be installed on any Windows machine and run in guest mode without any dependencies.
- `src/` — the entire VB6 application: the project file (`CROSS.vbp`), every
  form/module (`.frm`/`.frx`/`.bas`), the compiled dev build
  (`CROSS2.14.exe`), and the runtime asset folders the running app locates
  via `App.Path` at startup (`Ini/`, `Template/`, `Preview/`, `Output/`,
  `Temp/`). These all live together deliberately — `App.Path` resolves to
  wherever the running `.exe` physically sits, so the executable and the
  folders it looks up relative to itself can't be split apart without
  breaking those lookups.

## Setup

1. Open `src/CROSS.vbp` in the VB6 IDE.
2. Copy `src/DBCredentials.bas.example` to `src/DBCredentials.bas` and fill
   in real Oracle/FTP credentials (already excluded via `src/.gitignore` —
   never commit the filled-in copy).
3. Copy `src/Ini/Config.ini.example` to `src/Ini/Config.ini` and fill in the
   real server addresses/login (also gitignored).
4. Requires Golden Software Surfer 11 installed locally (see the `Reference=`
   line in `CROSS.vbp`) for the contour-mapping features to work. If
   contour-mapping is not required, Surfer is not mandatory.

## Note for anyone relying on the production deployment

This repo is a working copy shared for reference/review. If the relied-on database is well constructed, this repo can be deployed to any region beyond Guangdong, China.

## User Manual

A feature-by-feature walkthrough with screenshots.

### Contents

- [Overview](#overview)
- [Login](#login)
- [Time-scale queries: Hourly / Daily / Dekad / Monthly](#time-scale-queries-hourly--daily--dekad--monthly)
- [Single Period Query](#single-period-query)
- [Arbitrary Period Query](#arbitrary-period-query)
- [Arbitrary Period Decomposition Flowchart](#arbitrary-period-decomposition-flowchart)
- [Year-by-Year Period Query](#year-by-year-period-query)
- [Conditional Query](#conditional-query)
- [Season Classification](#season-classification)
- [Cold Air (Cold Surge) Statistics](#cold-air-cold-surge-statistics)
- [Station Selection](#station-selection)
- [Contour Mapping](#contour-mapping)
- [Charting: Histogram / Time Series / Scatter](#charting-histogram--time-series--scatter)
- [Import Data](#import-data)
- [Regional Statistics](#regional-statistics)
- [Surfer Settings](#surfer-settings)
- [User Management](#user-management)
- [User Activity Analysis](#user-activity-analysis)
- [Message Center](#message-center)
- [System Update](#system-update)
- [About](#about)

### Overview

A composite view of several modules in use at once, giving a sense of how
they compose in practice: a Dekad-period time series chart with a trend
line and equation, a histogram, a scatter plot with regression statistics,
a contour map of Guangdong with its color-level settings dialog open, a
station data table, and system-wide feature-usage pie/bar charts. Most of
these are independent, dockable child windows that a user can have open
side by side.

![Overview](src/Preview/Gallery/fig00.Overall.png)

### Login

The startup login dialog. Users sign in with an assigned account and
password, which determines their data-region scope and administrative
level; a Guest option is available for read-only access without an
account.

![Login](src/Preview/Gallery/fig01.login.png)

### Time-scale queries: Hourly / Daily / Dekad / Monthly

The core data-query screens, one per time scale — the toolbar's Hourly,
Daily, Dekad, Monthly, Seasonal, and Annual buttons all open the same style
of window at a different aggregation level. Each shows a station table on
the left and element/date-range pickers along the top, with results
exportable directly from the grid.

![Hourly Data Query](src/Preview/Gallery/fig02.hourly.png)
![Daily Data Query](src/Preview/Gallery/fig03.daily.png)
![Monthly Data Query](src/Preview/Gallery/fig04.monthly.png)

### Single Period Query

Computes aggregate statistics (average, max, min, sum, cumulative, etc.)
across stations over one user-defined start/end date range, rather than a
fixed calendar period. Useful for ad hoc questions like "what was the
average temperature during this specific two-week span."

![Single Period Query](src/Preview/Gallery/fig05.sngPeriod.png)

### Arbitrary Period Query

Extends the single-period idea to an arbitrary, potentially
non-contiguous, set of date ranges in one query — for example, comparing
the same festival week across the past 70 years at once, providing the
30-year normals, the anomaly, and the ranking (both from high to low and
low to high). This is the most important feature used in operational work
on a regular basis. The three screenshots below show the setup and results
side of the workflow.

![Arbitrary Period Query (1)](src/Preview/Gallery/fig06.ArbPeriod1.png)
![Arbitrary Period Query (2)](src/Preview/Gallery/fig07.ArbPeriod2.png)
![Arbitrary Period Query (3)](src/Preview/Gallery/fig08.ArbPeriod3.png)

### Arbitrary Period Decomposition Flowchart

Illustrates how an arbitrary, potentially non-contiguous period is
decomposed internally to optimize statistical efficiency — the logic
behind the Arbitrary Period Query feature above.

![Arbitrary Period Decomposition Flowchart](src/Preview/Gallery/fig24.Flowchart_ArbPeriod.png)

### Year-by-Year Period Query

Repeats the same custom period (e.g. "June 1–15") across a range of years
and lines the results up for year-over-year comparison, which is a common
requirement for climate trend analysis.

![Year-by-Year Period Query](src/Preview/Gallery/fig09.YerPeriod.png)

### Conditional Query

Queries and aggregates data with a threshold or condition applied — for
example, counting the number of days a given element exceeded (or fell
below) a value — across multiple elements in a single pass, rather than
requiring one query per element.

![Conditional Query](src/Preview/Gallery/fig10.Conditioal.png)

### Season Classification

Determines season onset/classification dates from the underlying
observation data and reports the associated statistics, supporting
climate-monitoring questions like "when did this year's summer onset
occur relative to the historical average."

![Season Classification](src/Preview/Gallery/fig11.Season.png)

### Cold Air (Cold Surge) Statistics

A dedicated statistics module for cold-air outbreak (cold surge) events,
distinct from the general time-scale queries because cold-surge
identification follows its own meteorological criteria.

![Cold Air Statistics](src/Preview/Gallery/fig12.CodAir.png)

### Station Selection

The shared station-picker dialog used across query modules: filters the
working station set by type (National vs. AWS), then narrows by zone,
city, county, or township, with separate list panes for national and AWS
stations (including hydrological and oceanic stations) so a large custom station set can be built up incrementally.

![Station Selection](src/Preview/Gallery/fig13.Station.png)

### Contour Mapping

Generates spatial contour maps of the selected element over 148 regions (including Guangdong) via the bundled Surfer integration, with controls for color-level
range/interval, a cool-warm or warm-cool color ramp, gridding method, and
grid resolution.

![Contour Mapping](src/Preview/Gallery/fig14.Contour.png)

### Charting: Histogram / Time Series / Scatter

The charting suite shared across query results: histograms with
distribution statistics (mean, quartiles, percentiles), time-series charts
with trend lines and regression equations, and scatter plots with
correlation/regression statistics — each with manual axis adjustment and
save-to-file export.

![Charting suite](src/Preview/Gallery/fig15.Chart.png)

### Import Data

Imports external CSV data into the system, with configurable attribute
column counts and header handling, then matches each row to station
metadata (township/county/city) and geocoded longitude/latitude — useful
for bringing in data collected outside CROSS.

![Import Data](src/Preview/Gallery/fig16.ImportData.png)

### Regional Statistics

Aggregates statistics over custom multi-city or multi-province regions
(rather than a single province), with ranking, multi-year average, and
anomaly columns, backed by its own region/station selection dialog for
building the custom region.

![Regional Statistics](src/Preview/Gallery/fig17.LargeRegion.png)

### Surfer Settings

Configures the Golden Software Surfer integration that powers Contour
Mapping: gridding method, grid size, output image resolution, and
plot-region presets at the province, city, county, or fully custom level.

![Surfer Settings](src/Preview/Gallery/fig18.SurferSet.png)

### User Management

Administrator screen for managing user accounts: work unit, administrative
region, data-access scope, permission level, and the set of IP addresses
authorized to download data under that account.

![User Management](src/Preview/Gallery/fig19.Users.png)

### User Activity Analysis

An audit view over query/download activity, filterable by department,
username, download IP, and date range, paired with feature-usage and
daily-active-IP summary charts — useful for understanding how the system
is actually being used across the province.

![User Activity Analysis](src/Preview/Gallery/fig20.UsersActivity.png)

### Message Center

A two-way feedback channel between end users and administrators (visible
here as a real support conversation about a map/shapefile issue), plus
system-wide broadcast notices such as "new long-time-series data added."

![Message Center](src/Preview/Gallery/fig21.MsgCenter.png)

### System Update

Checks the installed client version against the latest available and
displays release notes for the current or any prior version, supporting
the system's phased-rollout update model.

![System Update](src/Preview/Gallery/fig22.Update.png)

### About

Version and contact information alongside a plain-language summary of the
system's data sources, supported time/spatial scales, and statistics —
effectively CROSS's own one-paragraph elevator pitch.

![About](src/Preview/Gallery/fig23.About.png)
