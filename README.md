# TiBAO Purchase Intelligence v3.2

## SHJ filter design fix
- Brand, Make and Status remain multi-select.
- Their dropdowns now use the same compact checkbox/text design as the Dashboard filters.
- Fixed the oversized checkbox problem caused by SHJ toolbar CSS affecting nested dropdown inputs.

## SHJ performance improvement
- SHJ Replenishment is now paginated instead of rendering every filtered row at once.
- Default: 100 rows per page. Options: 50 / 100 / 250.
- KPI counts, filtered totals, Select All and exports still operate on the complete filtered result, not just the visible page.
- Previous / Next controls and page information are shown below the table.

All purchase/replenishment calculations remain unchanged.

## GitHub update
Replace index.html, app.js, engine.js and styles.css, commit, wait for Pages, then press Ctrl+F5. The site should show **Logic v3.2**.
