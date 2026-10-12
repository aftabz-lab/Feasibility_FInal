# GP Share equals Franchisee Commission %

1. Extract `Feasibility_GP_Share_Commission_Fix_GitHub_Upload.zip`.
2. Open **https://github.com/aftabz-lab/Feasibility_FInal** → **main** branch → repository root.
3. Use **Add file → Upload files** and replace only these four files:
   - `model.mjs`
   - `excel-exporter.js`
   - `app.js`
   - `index.html`
4. Commit the files and wait for the GitHub Pages deployment to complete successfully.
5. Open **https://aftabz-lab.github.io/Feasibility_FInal/** and press **Ctrl + Shift + R**.
6. Generate a fresh **Download Excel with Rules** file. Its formula is:

   `INFORMATION!B7 = 'AUTO GENERATED FEASIBILITY'!B18`

The dashboard GP Share uses that same Franchisee Commission percentage. The existing `Fr Com% vs GP Share%` table and commission calculation are retained. The original lookup selection is stored in row 28 of the existing hidden `Dashboard Rules` sheet so the B7/B18 link has no circular reference and the generated workbook can be loaded again.

All unrelated workbook cells, formulas, fonts, formatting, sheet order and visibility remain unchanged. Financial calculations, signatory rules and signs, map rules, nearby-outlet rules, catchment behavior and PDF layouts are preserved. The source workbook and other repository files stay in place.

Upload the four replacement files listed above. `SETUP.md` is the installation guide.
