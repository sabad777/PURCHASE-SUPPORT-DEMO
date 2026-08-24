# TiBAO Purchase Intelligence v2.2

GitHub Pages-ready browser application for analysing the Odoo Purchase Support Excel export.

## New in v2.2
- Multi-select **Brand** filter with checkboxes, search, Select All and Clear.
- Multi-select **Condition** filter (e.g. Critical + Reorder Soon together).
- Dedicated **Add Description** contains filter. Typing `water pump` finds descriptions such as `WATER PUMP W124`, `WATER PUMP GOLF 7`, etc.
- **Make** filter.
  - If the Odoo export later contains a Make/Vehicle Make column, it is used directly.
  - If not, the app derives known makes from the beginning of Internal Reference / old code, e.g. `VW 06L115562 FEBI` -> `VW`.
- Make is shown in the dashboard/planner and exported recommendation file.

## GitHub update
Replace these files in the existing repository:
- `index.html`
- `app.js`
- `engine.js`
- `styles.css`

GitHub Pages settings do not need to change. After committing, wait for deployment and press Ctrl+F5.
