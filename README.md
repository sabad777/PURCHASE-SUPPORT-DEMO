# TiBAO Purchase Intelligence — Web Prototype v2.3

Static browser-based proof-of-concept for the future Odoo Purchase Support module.

## What v2.3 adds
- Supports the latest Odoo Purchase Support export header layout (`Product ID`, `Internal Ref`, `Main OEM`, `Motorline`, `Bavaria`, `Tibao SHJ`, `All Company (Core)`, `Jan`-`Dec`, `Total Qty`, `Invoices`, `Last Sale`, `Received Qty`, `Receipts`, `Last Purchase`, and the 8 On Way / On Way 2 fields).
- Still accepts the previous longer header names for backward compatibility.
- Blank future months are safely ignored using the report date detected from the report title / Report Till Date field.
- Brand and Condition filters are multi-select.
- Add Description contains-search (for example `water pump`).
- Make filter derives vehicle make from the Internal Ref when there is no Odoo Make column.
- Incoming calculation prefers the three company-level On Way fields; Group totals are used only as fallback.
- Negative incoming values are protected from purchase calculations and flagged in Data Quality.

## Confirmed make-code mapping
- VW → VW
- BENZ → BENZ
- BMW → BMW
- ROV / LR → LAND ROVER
- TY → TOYOTA
- BE → BENTLEY
- P / PO → PORSCHE
- SKD → SKODA
- AD → AUDI
- UN → UNIVERSAL
- HY → HYUNDAI
- JE → JEEP

Additional aliases observed clearly in the current export are also recognized: AU → AUDI, JA → JAGUAR, HO → HONDA, SK → SKODA, OP → OPEL, MZ → MAZDA, MI → MITSUBISHI, MAS → MASERATI, WV → VW, JEEP → JEEP.

## GitHub Pages update
Replace these files in the existing repository root:
1. `index.html`
2. `engine.js`
3. `app.js`
4. `styles.css`

Commit the changes, wait for GitHub Pages to redeploy, then press `Ctrl + F5` and confirm the top badge says `Logic v2.3`.
