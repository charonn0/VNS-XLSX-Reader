# Test fixtures

Public XLSX files used for manual verification of the parser. Each one targets a specific feature or edge case.

| File | Source | Size | Sheets | Useful for |
|------|--------|-----:|--------|------------|
| `microsoft-financial-sample.xlsx` | [Microsoft Power BI demo data](https://go.microsoft.com/fwlink/?LinkID=521962) | 83 KB | 1 (`Sheet1`, ~700 rows) | real-world business workbook — currency formatting, dates, large row count, mixed text/numeric columns |
| `sheetjs-autofilter.xlsx` | [SheetJS test data](https://oss.sheetjs.com/test_files/AutoFilter.xlsx) | 58 KB | 11 (`No Filter`, `Just Filter`, `One Cond`, `Two Cond`, `Top10`, `Bot10`, `Average`, `NE`, `GT`, `AND Bounding`, `OR Range`) | many tabs in one workbook — exercises the tabbed-sheet UI |
| `sheetjs-pres.xlsx` | [sheetjs.com](https://sheetjs.com/pres.xlsx) | 9 KB | 1 | small clean smoke test — US presidents data |
| `excelize-sharedstrings.xlsx` | [qax-os/excelize](https://github.com/qax-os/excelize) test data | 7 KB | 2 | shared-strings table resolution |
| `excelize-calcchain.xlsx` | [qax-os/excelize](https://github.com/qax-os/excelize) test data | 6 KB | 1 | formula cells with cached values |
| `excelize-mergecell.xlsx` | [qax-os/excelize](https://github.com/qax-os/excelize) test data | 6 KB | 1 | merged-cell rectangles |

Each fixture is a public sample published by its upstream project. If a license question arises, delete and re-fetch from the URL above.
