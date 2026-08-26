# TiBAO Purchase Intelligence v2.5

GitHub Pages / browser-only purchase planning prototype.

## New in v2.5
- Projected **Current Stock-Out Date**
- Projected **Stock-Out Date With On Way + On Way 2**
- Optional **brand-specific lead time** in purchasing logic
- Default lead time plus a per-brand editable rules table
- Optional **MOQ** and **Order Multiple** rounding
- MOQ / multiple are OFF by default because true MOQ may be product-specific
- Detailed product explanation now shows lead time, expected new-order arrival, stock-out dates, MOQ/multiple and any quantity adjustment
- Existing manual Suggested Qty override and the 6-column Selected Purchase export remain unchanged
- No dead-stock value calculation was added

## GitHub update
Replace these four files in the existing repository:
- index.html
- app.js
- engine.js
- styles.css

Commit, wait for GitHub Pages to redeploy, then press Ctrl+F5. The page should show **Logic v2.5**.
