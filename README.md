# FUSO MEA Demand Planning — Executive Dashboard

**DSV Logistics × Mitsubishi FUSO | Jafza Free Zone, Dubai**

A live executive dashboard for the FUSO MEA spare parts distribution hub. Any change to the Excel model automatically rebuilds the dashboard via GitHub Actions and publishes to GitHub Pages.

---

## How It Works

```
Excel (data/) → push to GitHub → GitHub Actions → build.py → JSON → GitHub Pages → Dashboard
```

1. Edit `data/FUSO_Advanced_Model_v2.xlsx`
2. Commit and push to `main`
3. GitHub Actions detects the Excel change, runs `scripts/build.py`
4. Updated `data/dashboard_data.json` is committed automatically
5. Dashboard on GitHub Pages reflects new data within ~60 seconds

---

## Local Development

```bash
# 1. Install dependency
pip3 install openpyxl

# 2. Build JSON from Excel
python3 scripts/build.py

# 3. Serve locally (required — fetch() won't work on file://)
python3 -m http.server 8080

# 4. Open browser
open http://localhost:8080
```

---

## GitHub Setup (One-Time)

```bash
# From the fuso-demand-dashboard directory:
git init
git add .
git commit -m "feat: FUSO MEA executive demand planning dashboard"
git remote add origin https://github.com/YOUR_USERNAME/fuso-demand-dashboard.git
git push -u origin main
```

Then in GitHub → Settings → Pages → Source → `gh-pages` branch.

---

## Structure

```
fuso-demand-dashboard/
├── index.html                        # Executive dashboard (reads JSON)
├── data/
│   ├── FUSO_Advanced_Model_v2.xlsx  # Source of truth — edit this
│   └── dashboard_data.json          # Auto-generated — do not edit manually
├── scripts/
│   └── build.py                     # Excel → JSON pipeline
└── .github/
    └── workflows/
        └── update-dashboard.yml     # Auto-trigger on Excel push
```

---

## Dashboard Sections

| Section | Content |
|---|---|
| **KPIs** | Forecast accuracy, fill rate, SMOB%, portfolio value |
| **ABC-XYZ Matrix** | 9-cell segmentation with SKU counts and strategies |
| **Segmentation Forecasting** | WMA / Exp Smoothing / SBA / Causal / GPC Bridge |
| **Supply Network** | Japan / Chennai / Halberstadt lead times and SS |
| **SMOB Reduction** | 3-year AED recovery plan with disposition actions |
| **3-Year Roadmap** | Stabilise → Optimise → Excellence milestones |
| **SKU Portfolio** | Filterable table — all 50 SKUs with classification |

---

*Confidential — Interview Presentation Material — Vinayak Bhadani*
