# TiBAO Purchase Intelligence — Web Prototype

A static, browser-based proof-of-concept for the future Odoo Purchase Support module.

## What it does

- Reads the Odoo **Purchase Support Excel export** directly in the browser.
- Calculates purchase-support metrics using configurable rules.
- Keeps **All Company Qty**, **On Way**, and **On Way 2** separate.
- Matches the **same normalized OEM number in different brands**.
- Shows Critical, Reorder Soon, Wait Incoming, Fast/Medium/Slow Moving, Overstock, Dead Stock, and Suggested Purchase Qty.
- Filters the whole dashboard by Brand, Condition, Movement, Category, OEM/part search.
- Includes Brand Analysis and Data Quality views.
- Explains the calculation for every product.
- Exports filtered or selected suggestions back to Excel.

## Important design choice

Same-OEM other-brand stock is always *shown*. It is **not automatically deducted from the purchase recommendation by default**. The Settings page has an `Other-brand stock credit %` value. It starts at 0% until Purchasing + Management approve the business rule.

This is deliberate: a FEBI item and a TIBAO item may be technically interchangeable, but customer demand for the specific brand may be different.

## Expected Odoo fields

The parser is flexible, but the ideal export includes:

- Odoo Product ID
- Internal Reference
- Main OEM No Space
- Brand
- Brand Part No.
- Add Description
- Product Category
- Motorline Stock
- Bavaria Stock
- Tibao SHJ Stock
- All Company Qty (Core)
- Jan–Dec Ext Sales
- Total External Sales Qty
- External Sales Invoice Count
- Last External Sale Date
- Total External Purchase Qty
- External Purchase Count
- Last External Purchase Date
- Motorline On Way
- Bavaria On Way
- Tibao SHJ On Way
- Group On Way
- Motorline On Way 2
- Bavaria On Way 2
- Tibao On Way 2
- Group On Way 2
- Report Till Date / Data As Of

If `Group On Way` exists, it is used as the total. Otherwise the app sums the three company On Way columns. The same logic applies to On Way 2.

## GitHub Pages — easiest setup

1. Create a new GitHub repository.
2. Upload **all files in this folder** to the root of the repository.
3. In GitHub open **Settings → Pages**.
4. Under Build and deployment choose **Deploy from a branch**.
5. Select the `main` branch and `/ (root)` folder, then Save.
6. GitHub will show the Pages URL after deployment.

No Python or server installation is needed.

## Privacy note

This version is a static webpage. Excel parsing and calculations happen inside the user's browser; the application code does not upload the workbook to a backend.

However, **GitHub Pages itself is not an access-control system**. If the repository/page is public, anyone who knows the URL can open the empty application. They cannot see your Excel data unless they have the file, because the workbook is processed locally and is not stored by this app.

For a truly restricted company URL with individual logins, add real authentication later (company SSO, Cloudflare Access, Microsoft Entra, etc.) or host the same frontend inside the company network.

## Prototype logic

The current calculation rules are intentionally editable. They are for management validation, not yet a frozen purchasing policy.

Default approach:

- Demand rate = 40% overall monthly average + 60% recent 3-month average.
- Target = 3 months + 0.5 month safety stock.
- Critical projected cover = below 0.5 month.
- Reorder soon = below 1.5 months.
- Overstock = current stock above 6 months cover.
- Dead stock = stock exists, no report-period sales, and last sale is at least 12 months old.
- Other-brand credit = 0% by default.

Once management approves the Excel/web recommendations, these settings can be frozen and handed to the Odoo developer as the exact calculation specification.
