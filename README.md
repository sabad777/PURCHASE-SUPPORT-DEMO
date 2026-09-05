# TiBAO Purchase Intelligence v3.3

## Purchasing Brand Groups
You can now create approved brand families that share exact-same-OEM stock and incoming quantities in the Purchase Planner.

### TiBAO Family preset
Settings → Purchasing Brand Groups → **Add TiBAO Family**. It will select matching loaded brands among TIBAO, TIBAO HD, TIBAO EXTRA, TG AUTOTEILE, SHOWORLD and SHOWORLD NEW/SHOWORLD_NEW.

The preset starts at 100% credit for Stock, On Way and On Way 2. Review the brands and percentages, then press **Apply Settings**.

### Rules
- Exact normalized OEM match is mandatory.
- A brand can belong to only one enabled group.
- Brands outside a group use the existing global Other-brand stock credit setting.
- Dashboard Brand filters are only view filters; they do not create a purchasing group.

### Example
TIBAO stock 10 + TG AUTOTEILE On Way 2 60 + SHOWORLD On Way 2 40 = 110 effective same-OEM supply at 100% group credits. The TIBAO row is then assessed using that credited supply instead of ignoring the 100 pcs coming in the grouped brands.

## GitHub
Replace index.html, app.js, engine.js and styles.css. Commit, wait for Pages, then Ctrl+F5. Confirm **Logic v3.3**.
