import openpyxl, json, os, subprocess
from collections import defaultdict
from datetime import datetime, timedelta

# --- Config ---
xlsx_path = '/home/snkwok/Downloads/SKU List (19).xlsx'
repo_dir = '/home/snkwok/jerry-dashboard2'

# Compute month from yesterday (MTD for data), but date badge = today (download date)
yesterday = datetime.now() - timedelta(days=1)
month = yesterday.strftime('%Y-%m')
today_str = datetime.now().strftime('%Y-%m-%d')  # actual download date, not yesterday
print(f"Month: {month}, Date: {today_str}")

# === Load historical name lookup ===
full_json_path = os.path.join(repo_dir, 'sku_data_full.json')
with open(full_json_path, 'r') as f:
    full_data = json.load(f)
name_lookup = {}
for entry in full_data:
    sc = entry['sc']
    sn = entry.get('sn', '')
    if sn and sn.strip():
        name_lookup[sc] = sn
print(f'Historical SKUs with names: {len(name_lookup)}')

# === Parse XLSX ===
wb = openpyxl.load_workbook(xlsx_path)
ws = wb.active
headers = [cell.value for cell in ws[1]]
print(f"Headers: {headers}")

col_map = {}
for i, h in enumerate(headers):
    col_map[h] = i

sku_idx = col_map['primary_sku_code']
name_idx = col_map['primary_sku_name_chi']
gmv_idx = col_map['GMV']
qty_idx = col_map['Qty']

print(f"Column indices: sku={sku_idx}, name={name_idx}, gmv={gmv_idx}, qty={qty_idx}")

# Count empty names
empty_names = 0
total_rows = 0
rows = []
for row in ws.iter_rows(min_row=2, values_only=True):
    sc = row[sku_idx]
    if not sc:
        continue
    total_rows += 1
    gmv = float(row[gmv_idx] or 0)
    qty = int(row[qty_idx] or 0)
    sn_raw = row[name_idx]
    sn = (sn_raw or '').strip()
    if not sn:
        empty_names += 1
    rows.append({'sc': sc, 'sn': sn, 'gmv': gmv, 'qty': qty, 'm': month})

# Sort by GMV descending
rows.sort(key=lambda x: -x['gmv'])

# Fill empty names from history
filled_from_history = 0
still_empty = 0
for r in rows:
    if not r['sn']:
        if r['sc'] in name_lookup:
            r['sn'] = name_lookup[r['sc']]
            filled_from_history += 1
        else:
            still_empty += 1

print(f'Total: {len(rows)} SKUs, Empty in XLSX: {empty_names}/{total_rows}')
print(f'Filled from history: {filled_from_history}, Still empty: {still_empty}')
print(f'Total GMV: {sum(r["gmv"] for r in rows):.2f}')

# Generate compact JSON (no spaces between separators)
sku_data_json = json.dumps(rows, ensure_ascii=False, separators=(',', ':'))

# --- Update sku_data_full.json ---
old_count = len(full_data)
full_data = [e for e in full_data if e['m'] != month]
removed = old_count - len(full_data)
print(f'Removed {removed} old entries for {month}')

full_data.extend(rows)
full_data.sort(key=lambda x: (-int(x['m'].replace('-', '')), -x['gmv']))

with open(full_json_path, 'w') as f:
    json.dump(full_data, f, ensure_ascii=False)
print(f'sku_data_full.json: {len(full_data)} entries total')

# --- Update index.html ---
html_path = os.path.join(repo_dir, 'index.html')

with open(html_path, 'r') as f:
    content = f.read()

# 1. Replace inline skuData array
lines = content.split('\n')
sku_updated = False
sku_line = None
for i, line in enumerate(lines):
    stripped = line.strip()
    if stripped.startswith('const skuData = [') and stripped.endswith('];'):
        indent = line[:len(line) - len(line.lstrip())]
        lines[i] = indent + 'const skuData = ' + sku_data_json + ';'
        sku_line = i + 1
        sku_updated = True
        print(f'skuData updated at line {i+1} ({len(rows)} entries)')
        break

if not sku_updated:
    print('ERROR: skuData line not found!')
    exit(1)

content = '\n'.join(lines)

# Write text content FIRST
with open(html_path, 'w') as f:
    f.write(content)

# 2. Update date badge (re-read in binary mode)
with open(html_path, 'rb') as f:
    raw = f.read()

date_prefix = b'\xf0\x9f\x93\x85 \xe6\x95\xb8\xe6\x93\x9a\xe4\xb8\x8b\xe8\xbc\x89\xe6\x97\xa5\xe6\x9c\x9f: '
if date_prefix in raw:
    start = raw.find(date_prefix) + len(date_prefix)
    old_date = raw[start:start+10].decode('utf-8')
    raw = raw.replace(date_prefix + old_date.encode(), date_prefix + today_str.encode())
    print(f'Date badge: {old_date} \u2192 {today_str}')
else:
    print('WARNING: Date badge prefix not found in index.html')

with open(html_path, 'wb') as f:
    f.write(raw)

# --- Verify GP is NOT touched ---
with open(html_path, 'rb') as f:
    final_raw = f.read()
gp_check = b'"GP":8000000' in final_raw
sku_count = final_raw.count(b'"sc":')
print(f'GP:8000000 preserved: {gp_check}')
print(f'SKU count in file: {sku_count}')

# --- Final stats ---
print(f'\n=== SUMMARY ===')
print(f'XLSX: {len(rows)} SKUs parsed')
print(f'sku_data_full.json: {len(full_data)} entries total')
print(f'Date badge updated to {today_str}')
print(f'GP integrity: {"PASS" if gp_check else "FAIL"}')
