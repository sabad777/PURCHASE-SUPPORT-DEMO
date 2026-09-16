# TiBAO Purchase Intelligence v3.7

## Changes in V3.5

### 1) New condition: NO STOCK / NO PURCHASE — REVIEW
Replaces the old NO DEMAND condition.

Rule:
- Sales = 0
- All Company stock = 0
- On Way = 0
- On Way 2 = 0
- Purchase Qty = 0
- No credited same-OEM supply

Action: REVIEW ITEM / PURCHASE HISTORY.

This avoids assuming an item has no demand when the company never stocked or purchased it during the report period.

### 2) Exact-OEM Purchasing Brand Group duplicate-purchase protection
For every enabled Purchasing Brand Group, exact normalized OEM families are planned once.

- Combined smart demand is calculated across the group brands for that exact OEM.
- Existing stock/incoming is credited using the group's Stock / On Way / On Way 2 percentages.
- One family shortage is calculated.
- The shortage is allocated to the group's Preferred Purchase Brand when that brand exists for the OEM.
- Other group rows are automatically set to 0 suggested quantity and marked COVERED BY FAMILY PLAN, preventing duplicate buying.
- TiBAO Family preset defaults Preferred Purchase Brand to TIBAO when available.

### 3) Optional Brand Order Cycle
Settings now include:
- Use brand order cycle: On/Off
- Default order cycle days
- Brand-specific Order Cycle Days in Brand Purchasing Rules

When enabled, the planning horizon considers:
enabled Lead Time + enabled Order Cycle + Safety Cover.

This helps when you buy a supplier only every few weeks/months.

### GitHub update
Replace:
- index.html
- app.js
- engine.js
- styles.css

Commit, wait for GitHub Pages, then Ctrl+F5. Confirm **Logic v3.7**.


## v3.7 export update

- Purchase Planner → **Export Selected** now includes **Brand Number / Brand Part No.**
- SHJ Replenishment exports also use the clear **Brand Number** heading.
- Brand Number is read from the uploaded Odoo Excel field `Brand Part No.` / `Brand Part Number`.


## v3.7 — Group 1 filter

- Added a new **Group 1** multi-select filter to Purchase Planner.
- Group 1 is read dynamically from the uploaded Odoo Excel column named **Group 1**.
- Existing **Description** search/filter remains unchanged.
- Multiple Group 1 values can be selected together.
- The filter works together with Brand, Condition, Make, Movement, Demand Pattern, Category and Description filters.
- Demo data uses the demo category as Group 1 so the filter is visible before uploading a file.
