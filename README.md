# TiBAO Purchase Intelligence v3.0

## Purchase Planner
- Added Line / Serial # before the part number.
- Existing multi-brand selection remains available.
- Existing editable Suggested Purchase Qty and selected export remain unchanged.

## SHJ Replenishment
- Added Line / Serial # before the part number.
- Added Brand Part No.
- Column order now places Status and Action immediately after Brand Part No.
- Suggested From Motorline is editable before export.
- Brand filter is now multi-select with search, Select All and Clear.
- Added row checkboxes and Select All for the current filtered SHJ list.
- Added Export Selected and Clear Selection, same workflow as Purchase Planner.
- Changing SHJ filters removes selections that are no longer in the current filtered result.

## SHJ Excel export
The exported SHJ workbook contains only:
1. Line no
2. Old Number
3. OEM No Space
4. Add Description
5. Brand No
6. Brand
7. Suggested From Motorline
8. Motorline Stock
9. Bavaria Nearby Stock
10. SHJ Stock
11. Status

The edited Suggested From Motorline quantity is used in the export.

## GitHub
Replace:
- index.html
- app.js
- engine.js
- styles.css

Commit the changes, wait for GitHub Pages, then press Ctrl+F5.
The site should display Logic v3.0.
