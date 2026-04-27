# Niyu Excel Connector for Odoo 19

A true Excel connector for Odoo: build reusable profiles, export clean Excel files, download import templates, and import Excel files back into Odoo with validation and logs.

## Core idea
- No Microsoft OAuth
- No Excel Online dependency
- Simple business UI
- Profiles that normal users can run

## Main features
- One-click Excel export
- Downloadable Excel template from profile columns
- Excel import with dry-run mode
- Update, create, or upsert mode
- Scheduled export
- Error file for failed import rows
- Business presets for common Odoo models

## Optional server dependency for imports
Excel export works with xlsxwriter. Excel import requires `openpyxl` to be available on the Odoo server.
