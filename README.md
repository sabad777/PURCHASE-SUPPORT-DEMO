# TiBAO Purchase Intelligence v2.7

Bug-fix release for Purchase Planner selection.

## Fixed
- Clear Selection is always clickable.
- Clear Selection shows the current selected count.
- Selections remain restricted to the current filtered result.
- Select All selects the complete current filtered result.
- Added cache-busting (`?v=2.7`) to CSS and JavaScript references so GitHub Pages/browser caching cannot mix old and new versions.

## GitHub update
Replace all four files:
- index.html
- app.js
- engine.js
- styles.css

Commit, wait for GitHub Pages to redeploy, then press Ctrl+F5.
The page should show **Logic v2.7**.
