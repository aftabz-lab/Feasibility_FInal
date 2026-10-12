# Franchisee Commission = GP Share

1. Extract this ZIP.
2. Open https://github.com/aftabz-lab/Feasibility_FInal on the main branch.
3. Upload and replace ONLY these four files in the repository root: model.mjs, excel-exporter.js, app.js, index.html.
4. Commit the changes and wait for the GitHub Pages deployment to finish.
5. Open the dashboard and press Ctrl+Shift+R. Generate new Excel and PDF downloads; previously downloaded files retain their old rules.

GP Share is the input, with the previous 40% franchise default. Franchisee Commission % uses that same rate in the dashboard, Download 3-page PDF, Management PDF, and downloadable Excel.

Excel rules:
- INFORMATION!B7 contains the selected GP Share input.
- AUTO GENERATED FEASIBILITY!B18 = INFORMATION!B7.
- AUTO GENERATED FEASIBILITY!B60 retains its existing link to B18.

The prior Excel #NAME? compatibility fix and Commercial Hub / M reference fallback are retained. Other calculation formulas, fonts, formats, page layouts, signatory criteria, signature assets, map calculations, and catchment handling are unchanged. Existing OWN behavior retains a zero commission.

The PDF exporter and stylesheets are not included because their code is unchanged. PDF results receive the corrected rate from the shared model.
