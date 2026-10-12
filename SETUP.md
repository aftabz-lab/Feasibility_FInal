# Forecast references update

1. Extract this ZIP.
2. Open https://github.com/aftabz-lab/Feasibility_FInal on the main branch.
3. Upload and replace these five files in the repository root: app.js, model.mjs, excel-exporter.js, pdf-exporter.js, index.html. Do not delete or replace any other files or folders.
4. Commit the changes, wait for the Pages deployment to finish, then refresh the dashboard with Ctrl+Shift+R.

Sales Forecasting now shows Sales (Reference), FF (Reference)/Day, Basket (Reference), Profit (Reference), and Projected Basket Size (Reference) directly below Weighted score card.

Valid reference lookups remain unchanged. Missing/#N/A location-density-division lookups use the lowest Commercial Hub / M reference within the selected Division. The existing monthly-to-daily conversions and basket calculation remain in place. A valid projected-basket lookup or manual override is retained. A missing Division reference is not replaced with data from another Division.

The Excel export includes the corresponding fallback formulas, with helper calculations in the existing hidden Dashboard Rules sheet. PDF reference values use the same resolved figures; layouts, fonts, signatures, and signatory rules are preserved. The previous GP Share = Franchisee Commission fix is retained.

This ZIP contains only the five changed runtime files and these instructions. It does not include or replace source workbooks, stylesheets, signatures, catchment files, map configuration, or other assets.
