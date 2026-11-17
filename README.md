# Solar Camp Planning – MGVCL Bodeli Division

SDN entry page for PM Surya Ghar Solar Registration & Vendor Selection camps (18-11-2026 to 30-11-2026).

- `index.html` – entry page hosted on GitHub Pages (SDN login, date-wise camp entry, Division summary & Excel download)
- `Code.gs` – Google Apps Script backend; data is saved in the Google Sheet (`Entries` sheet), PINs in `Users` sheet

## Setup
1. Create a blank Google Sheet → Extensions → Apps Script → paste `Code.gs` → Save → run `setup` once (allow permissions).
2. Deploy → New deployment → Web app → Execute as: **Me**, Who has access: **Anyone** → Deploy → copy Web App URL.
3. In `index.html` set `const API_URL = "<Web App URL>";`
4. Upload `index.html` to a GitHub repository → Settings → Pages → Branch: `main` / root → Save.
5. Share `https://<github-username>.github.io/<repo-name>/` with SDNs along with their PIN from the `Users` sheet.
