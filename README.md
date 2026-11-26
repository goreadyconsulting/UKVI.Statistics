# UK Immigration Intelligence

Public, independent Home Office statistics site. No sign-in is required.

- The default homepage is the original Sponsored Study dashboard, using the existing read-only Supabase API and unchanged filter scope.
- `explore.html` provides the 20 official catalogue sections and interactive Entry Clearance Visa analysis.
- The daily Pages workflow retrieves the official catalogue and full detailed entry clearance workbook, verifies the expected headers and year/quarter structure, and builds year-partitioned static reporting data. An unsuccessful build does not deploy an incomplete snapshot.
- Source links for the other statistical domains are refreshed, but domain-specific charts and full database import are not implemented yet.
- The existing Supabase tables are not modified by the public sync. No privileged API keys or database passwords are needed or exposed. The original Sponsored Study dashboard still uses its existing dataset until parity and secure ingestion are established.
- Raw data is downloaded from the official source during the workflow and is not committed to GitHub; generated reporting data is published in the Pages artifact.
- Applications and outcomes measure separate events. The explorer requires a single applicant classification to avoid automatically mixing aggregate and detailed records.

Source: https://www.gov.uk/government/statistical-data-sets/immigration-system-statistics-data-tables
