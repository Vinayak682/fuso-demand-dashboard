#!/bin/bash
# ─────────────────────────────────────────────────────────────
#  FUSO MEA Dashboard Updater
#  Double-click this file to rebuild and push the dashboard
# ─────────────────────────────────────────────────────────────

# Change to the project folder
cd "$(dirname "$0")"

echo ""
echo "╔══════════════════════════════════════════╗"
echo "║   FUSO MEA DASHBOARD UPDATER             ║"
echo "║   DSV Logistics × Mitsubishi FUSO        ║"
echo "╚══════════════════════════════════════════╝"
echo ""

# Step 1 — Rebuild JSON from Excel
echo "► Step 1/3 — Reading Excel and rebuilding data..."
python3 scripts/build.py

if [ $? -ne 0 ]; then
  echo ""
  echo "✗ ERROR: build.py failed. Check your Excel file."
  echo ""
  read -p "Press Enter to close..."
  exit 1
fi

echo ""
echo "✓ Data rebuilt successfully"
echo ""

# Step 2 — Stage all changes
echo "► Step 2/3 — Staging changes..."
git add data/FUSO_Advanced_Model_v2.xlsx data/dashboard_data.json index.html scripts/build.py
echo "✓ Files staged"
echo ""

# Step 3 — Commit and push
echo "► Step 3/3 — Committing and pushing to GitHub..."
TIMESTAMP=$(date "+%Y-%m-%d %H:%M")
git commit -m "update: dashboard data refreshed ${TIMESTAMP}" 2>/dev/null

if git push 2>&1; then
  echo ""
  echo "╔══════════════════════════════════════════╗"
  echo "║   ✓ DASHBOARD UPDATED SUCCESSFULLY       ║"
  echo "║   GitHub Pages will refresh in ~60 sec   ║"
  echo "╚══════════════════════════════════════════╝"
else
  echo ""
  echo "⚠ Push failed — are you logged into GitHub?"
  echo "  Run in Terminal:  gh auth login"
fi

echo ""
read -p "Press Enter to close..."
