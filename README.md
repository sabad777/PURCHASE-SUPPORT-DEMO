# TiBAO Purchase Intelligence v2.4

GitHub Pages browser-only purchase support prototype.

## v2.4 changes
- Purchase Planner **Suggested Qty is editable**. The calculated suggestion is loaded as the default; purchaser can change it before export.
- Suggested Qty and Purchase Action are moved immediately after Brand Part No.
- `Export Selected` now creates a clean purchase Excel containing only:
  1. Old Number
  2. OEM No Space
  3. Add Description
  4. Available Quantity
  5. Total Incoming Quantity (On Way + On Way 2)
  6. Suggested Quantity (including purchaser edits)
- `Export Filtered` / top `Export Suggestions` still keep the detailed analytical export.

## GitHub update
Replace `index.html`, `app.js`, `engine.js`, and `styles.css` in the existing repository, commit, wait for GitHub Pages to deploy, then hard-refresh with Ctrl+F5.
