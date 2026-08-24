# TiBAO Purchase Intelligence — v2.1

Browser-only proof-of-concept for the future Odoo Purchase Support module.

## What changed in v2

- Month-by-month sales intelligence (not only total sales).
- Overall, recent 3-month and recent 6-month demand.
- Demand pattern: Regular, Rising, Falling, Intermittent, One-Time Spike, Dormant, Dead Stock.
- Active sales months and demand confidence.
- One-time spike protection: unusual single-month demand is sent for manual review instead of automatic purchasing.
- Purchase history is visible and analysed: Purchase Qty, Purchase Count, Last Purchase, Purchase/Sales ratio and warning signal.
- Dead Stock works from last-sale inactivity even if the item had earlier sales in the report period.
- Condition dropdown always includes Dead Stock and all other conditions, even when the current file has zero matching rows.
- All Company Qty, On Way and On Way 2 remain separate.
- Exact same OEM in other brands is checked automatically. Supersession is intentionally not used.
- Brand dashboard and Brand Analysis recalculate for the selected brand.
- Product detail popup explains monthly sales, purchase history, stock cover and the purchase calculation.
- Export includes all decision signals and month-by-month sales.

## Important calculation rule

Historical purchase quantity is **not subtracted again** when calculating suggested purchase quantity. Current stock already reflects received purchases. Purchase history is used as an intelligence / warning signal so the system can identify overbuying or unusual replenishment patterns.

Default **Proactive** mode recommends topping up toward Target Cover + Safety whenever projected cover is below the target. A traditional **Min/Max** mode is also available in Settings.

## Update the existing GitHub repository

Replace these four files in the repository root:

- `index.html`
- `engine.js`
- `app.js`
- `styles.css`

`.nojekyll` can stay unchanged. `README.md` is optional to replace.

After committing the replacements, GitHub Pages normally republishes automatically. Refresh the live page after a short wait. If the old interface is cached, press **Ctrl + F5**.

## Privacy

Excel parsing and calculations run in the browser. This static prototype does not send the uploaded workbook to an application server.


## v2.1 dead-stock fix
Never-sold stock is now eligible for Dead Stock. If Last Sale Date is blank, the engine uses Last Purchase Date when available; if neither date exists, it uses the consecutive zero-sales months in the report. This prevents never-sold stock from being stuck in NO SALES / REVIEW forever.
