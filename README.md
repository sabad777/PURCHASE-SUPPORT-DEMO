# TiBAO Purchase Intelligence v2.9

## New: SHJ Replenishment
A separate stock-balancing view has been added for Sharjah using the current Odoo Excel.

Because the current export does **not** contain separate Sharjah monthly sales, V2.9 temporarily estimates SHJ demand from overall smart demand. The final Odoo version should replace that assumption with actual SHJ sales history.

### Current logic
- Uses Tibao SHJ stock directly.
- Treats Bavaria stock as nearby backup when enabled.
- Estimates Sharjah demand as a configurable percentage of overall smart demand.
- Uses low target cover because the Motorline → Sharjah vehicle runs weekly.
- Keeps a configurable Motorline reserve before suggesting transfer.
- Avoids automatic transfers for dead/dormant/one-time/no-sales items or items below the minimum demand threshold.
- Shows Urgent Transfer, Transfer to SHJ, Use Bavaria Stock, Partial Transfer / Review, Purchase / Review, SHJ Stock OK and Review Demand.
- Exports the current SHJ replenishment list to Excel.

### Default prototype settings
- Estimated SHJ demand share: 25%
- Normal SHJ cover: 2 weeks
- Fast-mover SHJ cover: 3 weeks
- Motorline reserve: 1.5 months
- Bavaria nearby backup: ON
- Minimum group smart demand: 2 pcs/month

## GitHub update
Replace: index.html, app.js, engine.js and styles.css.
Commit, wait for GitHub Pages to redeploy, then press Ctrl+F5. The page should show **Logic v2.9**.
