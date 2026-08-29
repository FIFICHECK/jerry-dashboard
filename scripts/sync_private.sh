#!/bin/bash
# sync_private.sh — jerry-dashboard: copy index.html + data + reports to the PRIVATE repo + push
# B pilot (2026-08-29): 完整 dashboard 由 worker auth serve；gate 由 strip_gate 剝走
set -euo pipefail
SRC="${1:-/home/snkwok/jerry-dashboard}"
DST=/home/snkwok/dashboard-private-data/jerry-dashboard
cd "$SRC"
mkdir -p "$DST/data" "$DST/reports"
# 只喺 index.html 係完整 dashboard（冇 dashFrame marker）先同步 —— loader 唔可以冚 private
if ! grep -q 'dashFrame' index.html 2>/dev/null; then
  cp -f index.html /tmp/jerry_sync_index.html
  bash /home/snkwok/scripts/gate_strip_shared.sh /tmp/jerry_sync_index.html jerry-dashboard >/dev/null
  cp -f /tmp/jerry_sync_index.html "$DST/index.html"
fi
cp -f data/*.csv data/*.json data/*.js "$DST/data/" 2>/dev/null || true
cp -f sku_data_full.json "$DST/" 2>/dev/null || true
cp -rf reports/* "$DST/reports/" 2>/dev/null || true
cd "$DST"
if git status --short | grep -q .; then
  git add -A
  git -c user.email="hermes@fificheck.local" -c user.name="Hermes" commit -q -m "jerry-dashboard daily update $(date '+%F %T')"
  git push origin main -q
  echo "✅ private repo synced ($(date '+%F %T'))"
else
  echo "ℹ️ 冇嘢改 — skip push"
fi
