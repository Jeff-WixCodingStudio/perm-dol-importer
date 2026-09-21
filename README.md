# PERM DOL Importer

This GitHub Actions project downloads the newest official DOL PERM disclosure workbook, reduces it to chart aggregates, and publishes `public/perm-dashboard.json`.

It runs weekly on Monday at 16:00 UTC and can also be started manually from the repository's Actions page. It intentionally stores only aggregate chart data, not individual case records.

After this repository is created as a public GitHub repository, use the raw URL below in the Wix backend:

`https://raw.githubusercontent.com/OWNER/REPOSITORY/main/public/perm-dashboard.json`
