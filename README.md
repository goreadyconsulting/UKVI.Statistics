# UK Immigration Intelligence

Public, independent Home Office immigration analytics. No visitor accounts.

## Reports
- Sponsored Study remains the default: Main Applicant, Sponsored study, 2020 onwards, Issued + Refused decisions. Original country/period/quarter comparisons retained, with a new shared layout, quarterly volume chart and shareable global filters.
- Entry clearance explorer: full published dimensions, 2005 onwards, year-partitioned data, exact applicant classifications, presets, quarterly charts and filtered CSV.
- All 20 catalogue sections: automatically discovered official links. Domain-specific interactive analytics other than entry clearance and Sponsored Study are not yet implemented.

## Data
The scheduled GitHub Pages workflow downloads the official entry clearance workbook directly, validates the raw sheets, publishes static annual partitions and a small dedicated Sponsored Study snapshot. It checks annual Sponsored Study totals against the previous dashboard for 2020–2025 (5% or 150 count tolerance, whichever is larger). Unexpected schema changes or substantial historical mismatches block the new deployment for investigation. The existing Supabase snapshot remains a read-only fallback; it is not automatically updated by this workflow. No database write credentials are required.

Daily schedule: 06:37 UTC; manual workflow dispatch is available. Validated generated files are published both as a Pages artifact and a public repository snapshot. A raw GitHub read-only fallback is used if Pages has not served the generated files yet; unchanged releases do not cause daily snapshot commits. Source: https://www.gov.uk/government/statistical-data-sets/immigration-system-statistics-data-tables

The site is independent and not affiliated with the UK government. See each source for licensing and definitions.

## Wider immigration reports
The daily workflow downloads official ODS and XLSX files from each catalogue section, including linked regional and migration transparency pages. Published summary tables are plotted as individual values without inferred sums. Detailed data sheets with validated explicit date headers are partitioned by year; all non-axis categories are fixed to one published value, and ambiguous duplicate values are omitted instead of aggregated. Each source reports its own import status and has an original download link. Source-specific KPI definitions and specialist analytical models are not inferred automatically. Failed sources are marked unavailable without blocking the core Sponsored Study publication.
