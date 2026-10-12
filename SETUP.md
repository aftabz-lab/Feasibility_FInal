# Excel #NAME? fix

1. Extract this ZIP.
2. Open https://github.com/aftabz-lab/Feasibility_FInal on the main branch.
3. Upload and replace ONLY these three files in the repository root: excel-exporter.js, app.js, index.html. Keep every other file and folder.
4. Commit the upload and wait for the GitHub Pages deployment to finish.
5. Open the dashboard, press Ctrl+Shift+R, then click Download Excel with Rules to generate a new Excel file. Files downloaded before this update will still contain the old formulas.

This update replaces the newer reference fallback functions with compatible scalar Excel formulas. The five reference figures and their dependent projected footfall retain the same calculations. The lowest Commercial Hub / M reference is still selected within the chosen Division. Valid exact lookups and manual basket overrides remain unchanged.

The helper calculations stay in the existing hidden Dashboard Rules sheet. Existing worksheet layouts, formatting, formulas outside the corrected reference cells, GP Share / franchise commission rules, financial calculations, signatures, and PDFs are preserved.

The app.js and index.html changes only refresh the Excel export module in the browser cache. No PDF, model, stylesheet, signature, map, catchment, source workbook, or other asset is included in this ZIP.
