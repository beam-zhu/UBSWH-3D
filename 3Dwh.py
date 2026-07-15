# -*- coding: utf-8 -*-
import re
import os
import time
import datetime
import hashlib
import requests
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import json
from zoneinfo import ZoneInfo

# ========================================================
# 核心配置区
# ========================================================
PLAN_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vStLtLccpDSfrf_yU_T9WrNcZufN29BqkDsVS2r9ql1INK_61uA8UetkiW4DZ4_dv63o-DHzFz0tOAq/pub?gid=1276542862&single=true&output=csv"
ACTUAL_XLSX_URL = "https://docs.google.com/spreadsheets/d/1w1RvdGh_5LfIaxKHv0P-egK5kKzjLVZx/export?format=xlsx"
OUTPUT_HTML = "index.html"
TARGET_PASSWORD_HASH = "f0a36b9da192dc4732c232774766160f204bfe18be84c0a0dafce7040334b29f" 

# 🌟 核心修改：这里已经填入了你最新提供的 Web App URL
CONFIG_API_URL = "https://script.google.com/macros/s/AKfycbwRV4pnfbbNezMXCxR-CLCzb69oRXzzO7r8pZtZ6iPbcAyUngwhumuiMgjrAZJebDd7Kw/exec"

def get_deterministic_color(brand_name):
    hash_val = int(hashlib.md5(brand_name.encode('utf-8')).hexdigest(), 16)
    hue = hash_val % 360
    return f"hsl({hue}, 65%, 50%)"

GLOBAL_BRAND_COLORS = {
    'LINSY': '#D68F68', 'A区 (oversize沙发区)': '#7DA28A', 'B区 (沙发 Backup区)': '#6C8EA4',
    'G区不良品区': '#949BA2', 'Replica 区域': '#D4CBBE', 'MODE 椅子区': '#9E7E73',
    'LOFT 区': '#8B7AA3', 'Solidwood 区': '#C29B85', 'Boori区': '#C87284',
    'BOHOBOHO & Alpaka & Boori区': '#2C2D30', '补件区': '#8A5A58', 'loft & solidwood backup区': '#EEDCA5',
    '地面专属区': '#334155', '[当前空置]': '#E2E8F0', '[超过4品牌严重混放]': '#64748B'
}

def get_planned_info(zone, col, lvl, is_ground=False):
    if is_ground: return '地面专属区', GLOBAL_BRAND_COLORS['地面专属区']
    if zone in ['C', 'D', 'E', 'F', 'P']: return 'LINSY', GLOBAL_BRAND_COLORS['LINSY']
    elif zone == 'A': return 'A区 (oversize沙发区)', GLOBAL_BRAND_COLORS['A区 (oversize沙发区)']
    elif zone == 'B':
        if 2 <= lvl <= 5: return 'B区 (沙发 Backup区)', GLOBAL_BRAND_COLORS['B区 (沙发 Backup区)']
        return 'LINSY', GLOBAL_BRAND_COLORS['LINSY']
    elif zone == 'G': return 'G区不良品区', GLOBAL_BRAND_COLORS['G区不良品区'] 
    elif zone in ['H', 'J', 'K']:
        if zone == 'K' and (10 <= col <= 13) and (lvl in [1, 4, 5]): return 'Replica 区域', GLOBAL_BRAND_COLORS['Replica 区域']
        return 'LINSY', GLOBAL_BRAND_COLORS['LINSY']
    elif zone == 'L': return ('MODE 椅子区', GLOBAL_BRAND_COLORS['MODE 椅子区']) if lvl == 5 else ('LOFT 区', GLOBAL_BRAND_COLORS['LOFT 区'])
    elif zone == 'M': return 'LOFT 区', GLOBAL_BRAND_COLORS['LOFT 区']
    elif zone == 'Q':
        if lvl == 5: return 'MODE 椅子区', GLOBAL_BRAND_COLORS['MODE 椅子区']
        elif 2 <= lvl <= 4: return 'loft & solidwood backup区', GLOBAL_BRAND_COLORS['loft & solidwood backup区']
        return 'LOFT 区', GLOBAL_BRAND_COLORS['LOFT 区']
    elif zone == 'N': return 'Solidwood 区', GLOBAL_BRAND_COLORS['Solidwood 区']
    elif zone in ['T', 'W', 'X']:
        if zone == 'T' and col == 4: return '补件区', GLOBAL_BRAND_COLORS['补件区']
        if zone == 'X' and (col == 4 or (col == 3 and lvl == 1)): return 'Boori区', GLOBAL_BRAND_COLORS['Boori区']
        return 'Solidwood 区', GLOBAL_BRAND_COLORS['Solidwood 区']
    elif zone == 'Y':
        if col in [5, 6]: return 'BOHOBOHO & Alpaka & Boori区', GLOBAL_BRAND_COLORS['BOHOBOHO & Alpaka & Boori区']
        return 'Boori区', GLOBAL_BRAND_COLORS['Boori区']
    return '其他区域', '#8F9985'

COL_WIDTH, DEPTH, LVL_HEIGHT = 5.0, 3.0, 4.0
EASTERN_ALIGN_X, SOUTH_BASE_Y, NORTH_BASE_Y = 65.0, 95.0, 145.0
XWT_ZONE_COL_ORDER = {5: 0, 1: 1, 2: 2, 3: 3, 4: 4}
AISLE_GAP = 6.0
SHRINK = 0.92

def get_absolute_coords(zone, col, lvl):
    z = (lvl - 1) * LVL_HEIGHT
    if zone == 'Y': x, y = EASTERN_ALIGN_X - ((col - 1) * COL_WIDTH), NORTH_BASE_Y + 15
    elif zone in ['X', 'W', 'T']:
        step = XWT_ZONE_COL_ORDER.get(col, col - 1)
        x = EASTERN_ALIGN_X - (step * COL_WIDTH)
        y = NORTH_BASE_Y if zone == 'X' else (NORTH_BASE_Y - DEPTH if zone == 'W' else NORTH_BASE_Y - (DEPTH * 2) - 12)
    elif zone == 'H':
        x = 55.0 + 0.8
        y = SOUTH_BASE_Y - (col - 1) * COL_WIDTH - (COL_WIDTH / 2) if col <= 10 else SOUTH_BASE_Y + (col - 10) * COL_WIDTH - (COL_WIDTH / 2)
    elif zone == 'J':
        x = 55.0 - DEPTH - 0.8
        y = SOUTH_BASE_Y - (col - 1) * COL_WIDTH - (COL_WIDTH / 2) if col <= 11 else SOUTH_BASE_Y + (col - 11) * COL_WIDTH - (COL_WIDTH / 2)
    elif zone == 'K': x, y = 30.0 + 0.8, SOUTH_BASE_Y - (col - 1) * COL_WIDTH - (COL_WIDTH / 2)
    elif zone == 'L': x, y = 30.0 - DEPTH - 0.8, SOUTH_BASE_Y - (col - 1) * COL_WIDTH - (COL_WIDTH / 2)
    elif zone == 'P': x, y = 30.0 + 0.8, SOUTH_BASE_Y + (col - 1) * COL_WIDTH + (COL_WIDTH / 2)
    elif zone == 'Q': x, y = 30.0 - DEPTH - 0.8, SOUTH_BASE_Y + (col - 1) * COL_WIDTH + (COL_WIDTH / 2)
    elif zone == 'M':
        x = 5.0
        base_offset = (col - 1) * COL_WIDTH + (COL_WIDTH / 2)
        if col >= 3: base_offset += AISLE_GAP
        if col >= 6: base_offset += AISLE_GAP
        if col >= 9: base_offset += AISLE_GAP
        if col >= 12: base_offset += AISLE_GAP
        y = SOUTH_BASE_Y - base_offset
    elif zone == 'N':
        x = 5.0
        base_offset = (col - 1) * COL_WIDTH + (COL_WIDTH / 2) + AISLE_GAP
        if col >= 3: base_offset += AISLE_GAP
        y = SOUTH_BASE_Y + base_offset
    elif zone == 'A': x, y = 82 + (col - 1) * COL_WIDTH + (COL_WIDTH / 2), SOUTH_BASE_Y
    elif zone == 'B': x, y = 82 + (col - 1) * COL_WIDTH + (COL_WIDTH / 2), 80.0 - (DEPTH / 2)
    elif zone == 'C': x, y = 82 + (col - 1) * COL_WIDTH + (COL_WIDTH / 2), 80.0 - (DEPTH / 2) - DEPTH
    elif zone == 'D': x, y = 82 + (col - 1) * COL_WIDTH + (COL_WIDTH / 2), 56.0 - (DEPTH / 2)
    elif zone == 'E': x, y = 82 + (col - 1) * COL_WIDTH + (COL_WIDTH / 2), 56.0 - (DEPTH / 2) - DEPTH
    elif zone == 'F': x, y = 82 + (col - 1) * COL_WIDTH + (COL_WIDTH / 2), 32.0 - (DEPTH / 2)
    elif zone == 'G': x, y = 82 + (col - 1) * COL_WIDTH + (COL_WIDTH / 2), 32.0 - (DEPTH / 2) - DEPTH
    else: x, y = 200, 50
    return x, y, z

def is_valid_location(loc):
    loc = str(loc) if loc is not None else ''
    if not loc or loc.lower() in ['nan', 'none', '', 'nat']: return False
    if '$' in loc or ' ' in loc: return False
    if re.match(r'^[A-Z]+\d+-\d+$', loc): return True
    if re.match(r'^G[A-Z]+\d*$', loc): return True
    return False

def generate_html():
    print("📡 [1/4] 正在同步云端规划底座...")
    df_raw = pd.read_csv(PLAN_CSV_URL, header=None)
    all_raw_locs = []
    for col in df_raw.columns: all_raw_locs.extend(df_raw[col].dropna().astype(str).tolist())
    
    valid_locations = []
    for loc in all_raw_locs:
        loc = loc.strip()
        if not loc or loc == 'nan': continue
        match = re.match(r"([A-Z]+)(\d+)-(\d+)", loc)
        if match:
            zone, col_num, lvl_num = match.groups()
            valid_locations.append({'loc': loc, 'zone': zone, 'col': int(col_num), 'lvl': int(lvl_num), 'is_ground': False})
    df_locs = pd.DataFrame(valid_locations)

    print("📦 [2/4] 正在下载 Google Drive 最新 xlsx 文件... ")
    actual_db = {}
    temp_xlsx_path = "temp_inventory.xlsx"
    try:
        timestamp = int(time.time())
        url_with_timestamp = f"{ACTUAL_XLSX_URL}&t={timestamp}"
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'Cache-Control': 'no-cache, no-store, must-revalidate',
            'Pragma': 'no-cache',
            'Expires': '0'
        }
        response = requests.get(url_with_timestamp, headers=headers, timeout=60)
        response.raise_for_status()
        with open(temp_xlsx_path, 'wb') as f: f.write(response.content)
        with pd.ExcelFile(temp_xlsx_path) as xl_file:
            df_actual = xl_file.parse(sheet_name=0, skiprows=5, header=0)
        
        df_actual = df_actual[df_actual.iloc[:, 1].astype(str).apply(is_valid_location)]
        print(f"    XLSX 总行数（过滤后）: {len(df_actual)} ")
        
        for idx, row in df_actual.iterrows():
            if len(row) < 7: continue
            raw_loc = str(row.iloc[1]).strip()
            if not raw_loc or raw_loc.lower() in ['nan', 'none', '']: continue 
            raw_sku = str(row.iloc[2]).strip()
            clean_sku = raw_sku.split('~')[-1] if '~' in raw_sku else raw_sku
            brand = str(row.iloc[3]).strip() if len(row) > 3 else 'Other'
            try:
                qty = float(row.iloc[6]) if len(row) > 6 else 0.0
                if np.isnan(qty): qty = 0.0
            except: qty = 0.0
            
            if qty <= 0: continue
            
            if raw_loc not in actual_db: actual_db[raw_loc] = []
            actual_db[raw_loc].append({'sku': clean_sku, 'brand': brand, 'qty': int(qty)})
            if brand not in GLOBAL_BRAND_COLORS: GLOBAL_BRAND_COLORS[brand] = get_deterministic_color(brand)
    except Exception as e:
        print(f"⚠️ 下载 xlsx 失败: {e} ")
        if os.environ.get('GITHUB_ACTIONS') == 'true': raise e
    finally:
        try:
            if os.path.exists(temp_xlsx_path): os.remove(temp_xlsx_path)
        except: pass

    coords = [get_absolute_coords(z, c, l) for z, c, l in zip(df_locs['zone'], df_locs['col'], df_locs['lvl'])]
    df_locs['X'], df_locs['Y'], df_locs['Z'] = [c[0] for c in coords], [c[1] for c in coords], [c[2] for c in coords]

    ground_data = []
    ref_a = get_absolute_coords('A', 1, 1); ground_data.append({'loc': 'GA', 'X': ref_a[0] - 5.0, 'Y': ref_a[1], 'Z': 0, 'zone': 'GROUND', 'col': 1, 'lvl': 1, 'is_ground': True})
    ref_b = get_absolute_coords('B', 1, 1); ground_data.append({'loc': 'GBC', 'X': ref_b[0] - 5.0, 'Y': ref_b[1], 'Z': 0, 'zone': 'GROUND', 'col': 2, 'lvl': 1, 'is_ground': True})
    ref_d = get_absolute_coords('D', 1, 1); ground_data.append({'loc': 'GDE', 'X': ref_d[0] - 5.0, 'Y': ref_d[1], 'Z': 0, 'zone': 'GROUND', 'col': 3, 'lvl': 1, 'is_ground': True})
    ref_f = get_absolute_coords('F', 1, 1); ground_data.append({'loc': 'GFG', 'X': ref_f[0] - 5.0, 'Y': ref_f[1], 'Z': 0, 'zone': 'GROUND', 'col': 4, 'lvl': 1, 'is_ground': True})
    ref_j = get_absolute_coords('J', 11, 1); ground_data.append({'loc': 'GJ', 'X': ref_j[0], 'Y': ref_j[1] - 4.0, 'Z': 0, 'zone': 'GROUND', 'col': 5, 'lvl': 1, 'is_ground': True})
    ref_n4 = get_absolute_coords('N', 4, 1); ground_data.append({'loc': 'GN2', 'X': ref_n4[0], 'Y': ref_n4[1] + 4.0, 'Z': 0, 'zone': 'GROUND', 'col': 6, 'lvl': 1, 'is_ground': True})
    ref_n1 = get_absolute_coords('N', 1, 1); ref_m1 = get_absolute_coords('M', 1, 1); ground_data.append({'loc': 'GM1', 'X': ref_n1[0], 'Y': (ref_n1[1] + ref_m1[1])/2.0, 'Z': 0, 'zone': 'GROUND', 'col': 7, 'lvl': 1, 'is_ground': True})
    ref_n2 = get_absolute_coords('N', 2, 1); ref_n3 = get_absolute_coords('N', 3, 1); ground_data.append({'loc': 'GN1', 'X': ref_n2[0], 'Y': (ref_n2[1] + ref_n3[1])/2.0, 'Z': 0, 'zone': 'GROUND', 'col': 8, 'lvl': 1, 'is_ground': True})
    ref_m2 = get_absolute_coords('M', 2, 1); ref_m3 = get_absolute_coords('M', 3, 1); ground_data.append({'loc': 'GM2', 'X': ref_m2[0], 'Y': (ref_m2[1] + ref_m3[1])/2.0, 'Z': 0, 'zone': 'GROUND', 'col': 9, 'lvl': 1, 'is_ground': True})
    ref_m5 = get_absolute_coords('M', 5, 1); ref_m6 = get_absolute_coords('M', 6, 1); ground_data.append({'loc': 'GM5', 'X': ref_m5[0], 'Y': (ref_m5[1] + ref_m6[1])/2.0, 'Z': 0, 'zone': 'GROUND', 'col': 10, 'lvl': 1, 'is_ground': True})
    ref_m8 = get_absolute_coords('M', 7, 1); ref_m9 = get_absolute_coords('M', 9, 1); ground_data.append({'loc': 'GM8', 'X': ref_m8[0], 'Y': (ref_m8[1] + ref_m9[1])/2.0, 'Z': 0, 'zone': 'GROUND', 'col': 11, 'lvl': 1, 'is_ground': True})
    ref_m11 = get_absolute_coords('M', 11, 1); ref_m12 = get_absolute_coords('M', 12, 1); ground_data.append({'loc': 'GM11', 'X': ref_m11[0], 'Y': (ref_m11[1] + ref_m12[1])/2.0, 'Z': 0, 'zone': 'GROUND', 'col': 12, 'lvl': 1, 'is_ground': True})
    ref_t = get_absolute_coords('T', 1, 1); ground_data.append({'loc': 'GT', 'X': ref_t[0] + 5.0, 'Y': ref_t[1], 'Z': 0, 'zone': 'GROUND', 'col': 13, 'lvl': 1, 'is_ground': True})
    ref_w = get_absolute_coords('W', 1, 1); ground_data.append({'loc': 'GXW', 'X': ref_w[0] + 5.0, 'Y': ref_w[1], 'Z': 0, 'zone': 'GROUND', 'col': 14, 'lvl': 1, 'is_ground': True})

    ref_l14 = get_absolute_coords('L', 14, 1)
    ground_data.append({'loc': 'GL', 'X': ref_l14[0], 'Y': ref_l14[1] - 4.0, 'Z': 0, 'zone': 'GROUND', 'col': 99, 'lvl': 1, 'is_ground': True})
    ref_m13 = get_absolute_coords('M', 13, 1)
    ground_data.append({'loc': 'GM14', 'X': ref_m13[0], 'Y': ref_m13[1] - 4.0, 'Z': 0, 'zone': 'GROUND', 'col': 99, 'lvl': 1, 'is_ground': True})

    df_ground = pd.DataFrame(ground_data)
    df_locs = pd.concat([df_locs, df_ground], ignore_index=True)
    res = [get_planned_info(z, c, l, ig) for z, c, l, ig in zip(df_locs['zone'], df_locs['col'], df_locs['lvl'], df_locs['is_ground'])]
    df_locs['brand'], df_locs['color'] = [r[0] for r in res], [r[1] for r in res]

    # 🌟 核心修改：通过 API 实时读取 JSON，彻底抛弃有缓存的 CSV
    cloud_runtime_config = None
    cloud_cell_override_db = None
    cloud_actual_colors = None
    try:
        print("☁️ 正在通过 API 实时同步云端配置... ")
        api_res = requests.get(CONFIG_API_URL + '?t=' + str(int(time.time())), timeout=15)
        if api_res.status_code == 200:
            cloud_data = api_res.json()
            cloud_runtime_config = cloud_data.get('runtime_config')
            cloud_cell_override_db = cloud_data.get('cell_override_db')
            cloud_actual_colors = cloud_data.get('actual_brand_colors')
            print("✅ 云端配置(API)实时同步成功！ ")
        else:
            print(f"⚠️ API 请求失败，状态码: {api_res.status_code}，将使用默认配置。")
    except Exception as e:
        print(f"⚠️ 读取云端配置失败: {e}，将使用默认配置。")

    print("🧮 [3/4] 正在生成 3D 桥接数据与画布... ")
    python_to_js_cache = []
    for idx, row in df_locs.iterrows():
        locID = row['loc']
        items_in_bin = actual_db.get(locID, [])
        brands_in_bin = list(set([it['brand'] for it in items_in_bin if it['brand']]))
        brand_count = len(brands_in_bin)
        slices_data = []
        
        if brand_count == 0: slices_data.append({"brand": "[当前空置]", "color": GLOBAL_BRAND_COLORS['[当前空置]'], "items": []})
        elif brand_count > 4: slices_data.append({"brand": "[超过4品牌严重混放]", "color": GLOBAL_BRAND_COLORS['[超过4品牌严重混放]'], "items": [{"sku": it['sku'], "qty": it['qty'], "brand": it['brand']} for it in items_in_bin]})
        else:
            for b_name in brands_in_bin:
                b_color = GLOBAL_BRAND_COLORS.get(b_name, "#CBD5E1") 
                slices_data.append({"brand": b_name, "color": b_color, "items": [{"sku": it['sku'], "qty": it['qty']} for it in items_in_bin if it['brand'] == b_name]})
                
        h = LVL_HEIGHT * 0.92
        cz = row['Z']
        python_to_js_cache.append({
            "loc": str(locID), "zone": str(row['zone']), "col": int(row['col']), "lvl": int(row['lvl']),
            "is_ground": bool(row['is_ground']), "native_brand": str(row['brand']), "native_color": str(row['color']),
            "slices": slices_data, "orig_z": [cz, cz, cz, cz, cz+h, cz+h, cz+h, cz+h]
        })

    actual_total_stats = {}
    actual_total_qty = 0
    occupied_locations = 0
    
    for locID, items in actual_db.items():
        has_stock = False
        for it in items:
            brand = str(it.get('brand', '')).strip()
            qty = int(it.get('qty', 0))
            if brand and brand not in ['[当前空置]', '[超过4品牌严重混放]']:
                actual_total_stats[brand] = actual_total_stats.get(brand, 0) + qty
                actual_total_qty += qty
                if qty > 0: has_stock = True
        if has_stock:
            occupied_locations += 1

    total_locations = len(python_to_js_cache)
    occupancy_rate = round((occupied_locations / total_locations * 100), 1) if total_locations > 0 else 0.0
    print(f"   📦 仓库全量库存: {actual_total_qty} 件 | 📍 库位占用: {occupied_locations}/{total_locations} ({occupancy_rate}%)")

    fig = go.Figure()
    min_x, max_x = df_locs['X'].min() - 25, df_locs['X'].max() + 25
    min_y, max_y = df_locs['Y'].min() - 25, df_locs['Y'].max() + 25
    fig.add_trace(go.Mesh3d(x=[min_x, max_x, max_x, min_x, min_x, max_x, max_x, min_x], y=[min_y, min_y, max_y, max_y, min_y, min_y, max_y, max_y], z=[-1.5]*4 + [0.0]*4, i=[7,0,0,0,4,4,1,1,2,2,3,3], j=[0,1,2,3,5,7,2,5,3,6,0,7], k=[4,4,3,4,6,5,5,6,6,7,7,4], color='#CBD5E1', opacity=1.0, showlegend=False, hoverinfo='skip'))

    border_x, border_y, border_z = [], [], []
    for idx, row in df_locs.iterrows():
        is_v = row['zone'] in ['H', 'J', 'K', 'L', 'M', 'N', 'P', 'Q']
        if row['is_ground']: cw, cd, h = 2.0 * SHRINK, 2.0 * SHRINK, LVL_HEIGHT * 0.92
        else:
            cw = (DEPTH * 0.47 * SHRINK) if is_v else (COL_WIDTH * 0.47 * SHRINK)
            cd = (COL_WIDTH * 0.47 * SHRINK) if is_v else (DEPTH * 0.47 * SHRINK)
            h = LVL_HEIGHT * 0.92
        cx, cy, cz = row['X'], row['Y'], row['Z']
        x0, x1, y0, y1, z0, z1 = cx-cw, cx+cw, cy-cd, cy+cd, cz, cz+h
        edges = [(x0,y0,z0,x1,y0,z0),(x1,y0,z0,x1,y1,z0),(x1,y1,z0,x0,y1,z0),(x0,y1,z0,x0,y0,z0),(x0,y0,z1,x1,y0,z1),(x1,y0,z1,x1,y1,z1),(x1,y1,z1,x0,y1,z1),(x0,y1,z1,x0,y0,z1),(x0,y0,z0,x0,y0,z1),(x1,y0,z0,x1,y0,z1),(x1,y1,z0,x1,y1,z1),(x0,y1,z0,x0,y1,z1)]
        for e in edges: border_x.extend([e[0], e[3], None]); border_y.extend([e[1], e[4], None]); border_z.extend([e[2], e[5], None])
    fig.add_trace(go.Scatter3d(x=border_x, y=border_y, z=border_z, mode='lines', line=dict(color='#1E293B', width=2), hoverinfo='skip', showlegend=False, name='_BORDERS_'))

    for idx, row in df_locs.iterrows():
        is_v = row['zone'] in ['H', 'J', 'K', 'L', 'M', 'N', 'P', 'Q']
        if row['is_ground']: cw, cd, h = 2.0 * SHRINK, 2.0 * SHRINK, LVL_HEIGHT * 0.92
        else:
            cw = (DEPTH * 0.47 * SHRINK) if is_v else (COL_WIDTH * 0.47 * SHRINK)
            cd = (COL_WIDTH * 0.47 * SHRINK) if is_v else (DEPTH * 0.47 * SHRINK)
            h = LVL_HEIGHT * 0.92
        cx, cy, cz = row['X'], row['Y'], row['Z']
        x = [cx-cw, cx+cw, cx+cw, cx-cw, cx-cw, cx+cw, cx+cw, cx-cw]
        y = [cy-cd, cy-cd, cy+cd, cy+cd, cy-cd, cy-cd, cy+cd, cy+cd]
        z = [cz, cz, cz, cz, cz+h, cz+h, cz+h, cz+h]
        
        fig.add_trace(go.Mesh3d(x=x, y=y, z=z, name='_SHELF_CUBE_', i=[7,0,0,0,4,4,1,1,2,2,3,3], j=[0,1,2,3,5,7,2,5,3,6,0,7], k=[4,4,3,4,6,5,5,6,6,7,7,4], color=row['color'], customdata=[row['loc']]*8, text=[f"库位: {row['loc']}"]*8, lighting=dict(ambient=0.7, diffuse=0.85, specular=0.05), flatshading=True, hoverinfo='text', showlegend=False))
        
        hover_cx = cx + 0.1 if row['zone'] in ['H', 'K', 'P'] else (cx - 0.1 if row['zone'] in ['J', 'L', 'Q'] else cx)
        hover_cw = cw + 0.1
        hover_x = [hover_cx-hover_cw, hover_cx+hover_cw, hover_cx+hover_cw, hover_cx-hover_cw, hover_cx-hover_cw, hover_cx+hover_cw, hover_cx+hover_cw, hover_cx-hover_cw]
        hover_y = [cy-cd, cy-cd, cy+cd, cy+cd, cy-cd, cy-cd, cy+cd, cy+cd]
        hover_z = [cz, cz, cz, cz, cz+h, cz+h, cz+h, cz+h]
        hover_text = f"库位: {row['loc']} (地面)" if row['is_ground'] else f"库位: {row['loc']} <br>区域: {row['zone']}区 <br>品牌: {row['brand']}"
        
        fig.add_trace(go.Mesh3d(x=hover_x, y=hover_y, z=hover_z, name=f'_HOVER_{row["loc"]}', i=[7,0,0,0,4,4,1,1,2,2,3,3], j=[0,1,2,3,5,7,2,5,3,6,0,7], k=[4,4,3,4,6,5,5,6,6,7,7,4], color='white', opacity=0.01, hoverinfo='text', text=[hover_text] * 8, showlegend=False))
        
        if row['is_ground']:
            fig.add_trace(go.Scatter3d(x=[row['X']], y=[row['Y']], z=[row['Z'] + h + 0.5], mode='text', text=[row['loc']], textposition="top center", textfont=dict(size=12, color="#0F172A", family="Arial Black"), showlegend=False, hoverinfo='skip'))

    ZONE_Z_OFFSET_MAP = {'A': 2.0, 'B': 16.0, 'C': 2.0, 'D': 11.0, 'E': 2.0, 'F': 7.0, 'G': 2.0, 'H': 2.0, 'J': 6.0, 'K': 10.0, 'L': 2.0, 'M': 6.0, 'N': 2.0, 'P': 6.0, 'Q': 2.0, 'T': 2.0, 'W': 6.0, 'X': 10.0, 'Y': 2.0}
    for zone, group in df_locs.groupby('zone'):
        if zone == 'GROUND': continue
        mean_x, mean_y = np.mean(group['X']), np.mean(group['Y'])
        target_text_z = group['Z'].max() + LVL_HEIGHT + ZONE_Z_OFFSET_MAP.get(zone, 2.0)
        fig.add_trace(go.Scatter3d(x=[mean_x], y=[mean_y], z=[target_text_z], mode='text', text=[f"{zone}区"], textposition="top center", textfont=dict(size=19, color="#0F172A"), showlegend=False, hoverinfo='skip'))

    fig.update_layout(autosize=True, margin=dict(l=0, r=0, b=0, t=0), paper_bgcolor='#F8FAFC', scene=dict(xaxis=dict(visible=False), yaxis=dict(visible=False), zaxis=dict(visible=False), aspectmode='data', camera=dict(projection=dict(type="orthographic"), eye=dict(x=-0.8, y=-0.8, z=3.5), up=dict(x=0, y=0, z=1)), hovermode='closest'), dragmode="turntable")
    html_content = fig.to_html(include_plotlyjs='cdn', config={'scrollZoom': True, 'responsive': True, 'displayModeBar': False, 'doubleClick': 'reset'})
    cache_buster_meta = '''<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no"> <meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate"> <meta http-equiv="Pragma" content="no-cache"> <meta http-equiv="Expires" content="0">'''
    html_content = html_content.replace('<head>', '<head>' + cache_buster_meta)

    print("🎛️ [4/4] 正在编译前端交互引擎... ")
    nz_now = datetime.datetime.now(ZoneInfo('Pacific/Auckland'))
    data_timestamp = nz_now.strftime('%Y-%m-%d %H:%M:%S')

    js_global_colors_string = json.dumps(GLOBAL_BRAND_COLORS)
    js_array_string = json.dumps(python_to_js_cache)
    js_actual_total_stats = json.dumps(actual_total_stats)
    js_actual_total_qty = json.dumps(actual_total_qty)
    js_total_locations = json.dumps(total_locations)
    js_occupied_locations = json.dumps(occupied_locations)
    js_occupancy_rate = json.dumps(occupancy_rate)

    default_config_list = [
        {"org_name": "LINSY", "color": "#D68F68", "label": "LINSY"}, 
        {"org_name": "A区 (oversize沙发区)", "color": "#7DA28A", "label": "A区 (oversize沙发区)"},
        {"org_name": "B区 (沙发 Backup区)", "color": "#6C8EA4", "label": "B区 (沙发 Backup区)"}, 
        {"org_name": "G区不良品区", "color": "#949BA2", "label": "G区不良品区"},
        {"org_name": "Replica 区域", "color": "#D4CBBE", "label": "Replica 区域"}, 
        {"org_name": "MODE 椅子区", "color": "#9E7E73", "label": "MODE 椅子区"},
        {"org_name": "LOFT 区", "color": "#8B7AA3", "label": "LOFT 区"}, 
        {"org_name": "Solidwood 区", "color": "#C29B85", "label": "Solidwood 区"},
        {"org_name": "Boori区", "color": "#C87284", "label": "Boori区"}, 
        {"org_name": "BOHOBOHO & Alpaka & Boori区", "color": "#2C2D30", "label": "BOHOBOHO & Alpaka & Boori区"},
        {"org_name": "补件区", "color": "#8A5A58", "label": "补件区"}, 
        {"org_name": "loft & solidwood backup区", "color": "#EEDCA5", "label": "loft & solidwood backup区"}
    ]
    
    final_runtime_config = cloud_runtime_config if cloud_runtime_config else default_config_list
    final_cell_override_db = cloud_cell_override_db if cloud_cell_override_db else {}  
    final_actual_colors = cloud_actual_colors if cloud_actual_colors else {}

    js_config_string = json.dumps(final_runtime_config)
    js_overrides_string = json.dumps(final_cell_override_db)
    js_actual_colors_string = json.dumps(final_actual_colors)
    js_api_url_string = json.dumps(CONFIG_API_URL)

    interactive_control_script = r'''
<style>
body { margin: 0; overflow: hidden; font-family: sans-serif; }
.switch-btn { flex: 1; padding: 8px; font-size: 11px; font-weight: bold; border: 2px solid #CBD5E1; cursor: pointer; transition: all 0.2s ease; text-align: center; border-radius: 6px; background: #FFFFFF; color: #64748B; }
.switch-btn.active { background: #10B981; color: white; border-color: #047857; }
.switch-btn.active-actual { background: #3B82F6; color: white; border-color: #1D4ED8; }
#super-legend-panel { transition: transform 0.3s ease; max-height: 80vh; }
#nav-toggle-btn { position: absolute; top: 10px; left: 10px; z-index: 10001; background: #3B82F6; color: white; border: none; border-radius: 8px; padding: 8px 12px; font-size: 14px; cursor: pointer; display: none; }
@media (max-width: 768px) {
#nav-toggle-btn { display: block; }
#super-legend-panel { width: 85vw !important; max-width: 300px !important; transform: translateX(-120%); opacity: 0; }
#super-legend-panel.nav-open { transform: translateX(0) !important; opacity: 1 !important; }
#data-timestamp-box { top: 10px !important; right: 10px !important; padding: 5px 8px !important; }
#data-timestamp { font-size: 10px !important; }
}
#control-panel { position: absolute; bottom: 20px; right: 20px; z-index: 10000; background: rgba(255,255,255,0.95); border-radius: 12px; padding: 8px; box-shadow: 0 4px 20px rgba(0,0,0,0.15); display: flex; flex-direction: column; gap: 6px; border: 1px solid #E2E8F0; }
.control-row { display: flex; gap: 6px; }
.ctrl-btn { flex: 1; padding: 10px 0; border: 1px solid #CBD5E1; background: white; border-radius: 8px; cursor: pointer; font-size: 16px; font-weight: bold; color: #475569; transition: all 0.2s; display: flex; justify-content: center; align-items: center; }
.ctrl-btn:active { transform: scale(0.95); }
.ctrl-btn.active { background: #3B82F6; color: white; border-color: #2563EB; }
@media (max-width: 768px) { #control-panel { bottom: 10px; right: 10px; padding: 6px; } .ctrl-btn { padding: 12px 0; font-size: 18px; } }
.locked { display: none !important; }
#pwd-modal { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 100002; justify-content: center; align-items: center; }
.pwd-box { background: white; padding: 30px; border-radius: 12px; max-width: 400px; width: 90%; text-align: center; }
</style>
<button id="nav-toggle-btn" onclick="toggleNav()">☰ 菜单</button>
<div id="data-timestamp-box" style="position: absolute; top: 20px; right: 20px; background: rgba(255,255,255,0.95); padding: 10px 14px; border-radius: 10px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); z-index: 9999; border: 1px solid #E2E8F0; text-align: center;">
<div id="data-update-label" style="font-size: 10px; color: #666;">📊 数据更新 (NZ Time)</div>
<div id="data-timestamp" style="font-size: 13px; font-weight: bold; color: #0F172A;">DATA_TIMESTAMP_PLACEHOLDER</div>
<div style="display:flex; gap:5px; align-items:center; justify-content:center; margin-top: 5px;">
<button onclick="toggleLanguage()" id="lang-toggle-btn" style="background:#F1F5F9; color:#0F172A; border:1px solid #CBD5E1; border-radius:4px; padding:3px 8px; font-size:11px; cursor:pointer; font-weight:bold;">EN/中</button>
<button onclick="forceRefreshData()" style="background:#3B82F6; color:white; border:none; border-radius:4px; padding:3px 8px; font-size:11px; cursor:pointer;">🔄 <span id="btn-refresh-text">刷新</span></button>
</div>
</div>
<div id="super-legend-panel" style="position: absolute; top: 20px; left: 20px; background: rgba(255,255,255,0.98); padding: 16px; border-radius: 12px; box-shadow: 0 6px 20px rgba(0,0,0,0.1); z-index: 9999; width: 380px; border: 1px solid #E2E8F0; max-height: 90vh; overflow-y: auto;">
<div style="background: #F1F5F9; padding: 4px; border-radius: 8px; display: flex; gap: 4px; margin-bottom: 12px;">
<div id="view-plan-btn" class="switch-btn active" onclick="switchGlobalView('PLAN')">🟢 规划</div>
<div id="view-actual-btn" class="switch-btn" onclick="switchGlobalView('ACTUAL')">🔵 实际</div>
</div>
<div style="border-bottom: 2px solid #F1F5F9; padding-bottom: 6px; margin-bottom: 10px; display:flex; justify-content:space-between; align-items:center;">
<h4 id="legend-panel-title" style="margin: 0; font-size: 13px;">📊 预期规划品牌图例</h4>
<button id="reset-master-btn" class="lockable" onclick="resetToDefault()" style="background:#EF4444; color:white; border:none; border-radius:4px; padding:2px 8px; font-size:10px; cursor:pointer;">恢复初始</button>
</div>

<div id="sku-search-box" style="background: #F8FAFC; padding: 8px; border-radius: 8px; border: 1px dashed #5B7B9C; margin-bottom: 12px; display: none;">
    <label style="font-size: 11px; display:block; margin-bottom:4px; font-weight: bold;">🔍 SKU 搜索:</label>
    <div style="display: flex; gap: 4px;">
        <input type="text" id="sku-search-input" placeholder="输入 SKU（如：1234）" style="flex:1; padding: 6px; border: 1px solid #CBD5E1; border-radius: 4px; font-size: 11px;">
        <button onclick="searchSKU()" style="background: #3B82F6; color: white; border: none; padding: 6px 12px; border-radius: 4px; font-size: 11px; font-weight: bold; cursor: pointer;">搜索</button>
    </div>
    <div id="sku-search-results" style="margin-top: 8px; max-height: 200px; overflow-y: auto;"></div>
</div>

<div id="planning-tools-box" style="background: #F8FAFC; padding: 8px; border-radius: 8px; border: 1px dashed #5B7B9C; margin-bottom: 12px;">
<label id="quick-tool-label" style="font-size: 11px; display:block; margin-bottom:4px; font-weight: bold;">📐 快速改色工具:</label>
<textarea id="target-loc" rows="2" placeholder="如：Q01-01~Q01-04" style="width: 100%; padding: 4px; border: 1px solid #CBD5E1; border-radius: 4px; box-sizing: border-box; font-size: 11px;"></textarea>
<div style="display: flex; gap: 4px; margin-top: 6px;">
<input type="text" id="new-brand" placeholder="品牌" style="flex:1; padding: 4px; border: 1px solid #CBD5E1; border-radius: 4px; font-size: 11px;">
<input type="color" id="new-color" value="#DF9F57" style="width: 24px; height: 22px; border: 1px solid #CBD5E1; border-radius: 4px;">
<button id="apply-btn" class="lockable" onclick="applyLocationChange()" style="background: #5B7B9C; color: white; border: none; padding: 4px 10px; border-radius: 4px; font-size: 11px; font-weight: bold; cursor: pointer;">修改</button>
</div>
</div>
<div id="legend-list" style="display: flex; flex-direction: column; gap: 8px;"></div>
<button id="add-brand-btn" class="lockable" onclick="addNewBrand()" style="width:100%; padding:8px; background:#10B981; color:white; border:none; border-radius:6px; font-size:12px; font-weight:bold; cursor:pointer; margin-top:10px;">➕ 增加规划品牌</button>
</div>
<div id="control-panel">
<div class="control-row">
<button class="ctrl-btn active" id="btn-rotate" onclick="setMode('turntable', this)">🔄</button>
<button class="ctrl-btn" id="btn-pan" onclick="setMode('pan', this)">✋</button>
<button class="ctrl-btn" onclick="showPwdModal()" style="background:#FEF3C7; color:#D97706; border-color:#FCD34D; font-size:14px;">🔒</button>
</div>
<div class="control-row">
<button class="ctrl-btn" onclick="zoomCamera(0.8)">➕</button>
<button class="ctrl-btn" onclick="zoomCamera(1.25)">➖</button>
<button class="ctrl-btn" onclick="resetCamera()">🏠</button>
</div>
</div>
<div id="pwd-modal">
<div class="pwd-box">
<h3 id="pwd-title">🔐 输入编辑密码</h3>
<input type="password" id="pwd-input" placeholder="请输入密码" onkeypress="if(event.key==='Enter') verifyPwd()" style="width: 80%; padding: 10px; border: 2px solid #E2E8F0; border-radius: 6px; margin-bottom: 15px;">
<div>
<button id="pwd-cancel-btn" onclick="closePwdModal()" style="padding: 8px 16px; background: #E2E8F0; border: none; border-radius: 6px; cursor: pointer; margin: 0 5px;">取消</button>
<button id="pwd-confirm-btn" onclick="verifyPwd()" style="padding: 8px 16px; background: #3B82F6; color: white; border: none; border-radius: 6px; cursor: pointer; margin: 0 5px;">确认</button>
</div>
</div>
</div>
<script>
let runtime_config = SERVER_CONFIG_INJECT_PLACEHOLDER || [];
let cell_override_db = SERVER_OVERRIDES_INJECT_PLACEHOLDER || {};
let actual_brand_colors = ACTUAL_COLORS_INJECT_PLACEHOLDER || {};
let actual_total_stats = ACTUAL_TOTAL_STATS_PLACEHOLDER || {};
let actual_total_qty = ACTUAL_TOTAL_QTY_PLACEHOLDER || 0;
let total_locations = TOTAL_LOCATIONS_PLACEHOLDER || 0;
let occupied_locations = OCCUPIED_LOCATIONS_PLACEHOLDER || 0;
let occupancy_rate = OCCUPANCY_RATE_PLACEHOLDER || 0;

const CONFIG_API_URL = CONFIG_API_URL_PLACEHOLDER;
let GLOBAL_CURRENT_VIEW = "PLAN";
let server_data_cache = SERVER_DATA_INJECT_PLACEHOLDER || [];
let GLOBAL_COLOR_POOL = SERVER_COLORS_INJECT_PLACEHOLDER || {};

Object.assign(GLOBAL_COLOR_POOL, actual_brand_colors);

try {
    let saved_actual = localStorage.getItem("warehouse_twin_actual_colors_2026");
    if (saved_actual) { 
        let parsed = JSON.parse(saved_actual); 
        Object.assign(actual_brand_colors, parsed); 
        Object.assign(GLOBAL_COLOR_POOL, parsed); 
    }
} catch(e) {}

const translations = {
    zh: {
        dataUpdate: "📊 数据更新 (NZ Time)", refresh: "刷新", confirmRefresh: "确定刷新？",
        plan: "🟢 规划", actual: "🔵 实际", planTitle: "📊 预期规划品牌图例", actualTitle: "🔍 实盘现存品牌清点 (全量)",
        reset: "恢复初始", confirmReset: "确定要恢复所有初始规划并清除本地和云端保存的修改吗？",
        quickTool: "📐 快速改色工具:", locPlaceholder: "如：Q01-01~Q01-04", brandPlaceholder: "品牌", apply: "修改",
        addBrand: "➕ 增加规划品牌", promptBrandName: "请输入新品牌名称：", promptBrandExists: "该品牌已存在于规划中！",
        promptBrandColor: "请输入品牌颜色 HEX 值 (如 #FF5733)，或留空使用默认蓝色：", 
        totalInventory: "📦 仓库总库存 (全量)", 
        occupancyRate: "📍 库位占用率",
        skuSearch: "🔍 SKU 搜索:", skuPlaceholder: "输入 SKU（如：1234）", search: "搜索",
        deleteConfirm: "删除", restoreConfirm: "恢复", pwdTitle: "🔐 输入编辑密码", pwdPlaceholder: "请输入密码",
        cancel: "取消", confirm: "确认", unlockAlert: "🔒 请先点击右下角 🔒 按钮输入密码解锁编辑功能！", wrongPwd: "密码错误！"
    },
    en: {
        dataUpdate: "📊 Data Update (NZ Time)", refresh: "Refresh", confirmRefresh: "Are you sure to refresh?",
        plan: "🟢 Plan", actual: "🔵 Actual", planTitle: "📊 Planned Brand Legend", actualTitle: "🔍 Actual Inventory (Full)",
        reset: "Reset", confirmReset: "Reset all plans and clear local/cloud changes?",
        quickTool: "📐 Quick Color Tool:", locPlaceholder: "e.g.: Q01-01~Q01-04", brandPlaceholder: "Brand", apply: "Apply",
        addBrand: "➕ Add Planned Brand", promptBrandName: "Enter new brand name:", promptBrandExists: "This brand already exists!",
        promptBrandColor: "Enter HEX color (e.g. #FF5733) or leave empty:", 
        totalInventory: "📦 Total Inventory (Full)", 
        occupancyRate: "📍 Bin Occupancy Rate",
        skuSearch: "🔍 SKU Search:", skuPlaceholder: "Enter SKU (e.g.: 1234)", search: "Search",
        deleteConfirm: "Delete", restoreConfirm: "Restore", pwdTitle: "🔐 Enter Password", pwdPlaceholder: "Enter password",
        cancel: "Cancel", confirm: "OK", unlockAlert: "🔒 Click the 🔒 button at bottom right to unlock!", wrongPwd: "Wrong password!"
    }
};
let currentLang = localStorage.getItem('warehouse_lang') || 'zh';
function t(key) { try { return (translations[currentLang] && translations[currentLang][key]) ? translations[currentLang][key] : key; } catch(e) { return key; } }

function applyLanguage() {
    document.getElementById('data-update-label').innerText = t('dataUpdate');
    document.getElementById('btn-refresh-text').innerText = t('refresh');
    document.getElementById('view-plan-btn').innerText = t('plan');
    document.getElementById('view-actual-btn').innerText = t('actual');
    document.getElementById('reset-master-btn').innerText = t('reset');
    document.getElementById('quick-tool-label').innerText = t('quickTool');
    document.getElementById('target-loc').placeholder = t('locPlaceholder');
    document.getElementById('new-brand').placeholder = t('brandPlaceholder');
    document.getElementById('apply-btn').innerText = t('apply');
    document.getElementById('add-brand-btn').innerText = t('addBrand');
    document.getElementById('pwd-title').innerText = t('pwdTitle');
    document.getElementById('pwd-input').placeholder = t('pwdPlaceholder');
    document.getElementById('pwd-cancel-btn').innerText = t('cancel');
    document.getElementById('pwd-confirm-btn').innerText = t('confirm');
    document.getElementById('lang-toggle-btn').innerText = currentLang === 'zh' ? 'EN/中' : '中/EN';
    
    const skuSearchBox = document.getElementById('sku-search-box');
    if (skuSearchBox) {
        if (GLOBAL_CURRENT_VIEW === 'ACTUAL') {
            skuSearchBox.style.display = 'block';
            document.querySelector('#sku-search-box label').innerText = t('skuSearch');
            document.getElementById('sku-search-input').placeholder = t('skuPlaceholder');
            document.querySelector('#sku-search-box button').innerText = t('search');
        } else {
            skuSearchBox.style.display = 'none';
        }
    }
    
    if(GLOBAL_CURRENT_VIEW === 'PLAN') document.getElementById("legend-panel-title").innerText = t('planTitle');
    else document.getElementById("legend-panel-title").innerText = t('actualTitle');
    if (typeof renderControlPanel === 'function') renderControlPanel();
}
function toggleLanguage() { currentLang = currentLang === 'zh' ? 'en' : 'zh'; localStorage.setItem('warehouse_lang', currentLang); applyLanguage(); }

function searchSKU() {
    const searchInput = document.getElementById('sku-search-input').value.trim().toLowerCase();
    const resultsDiv = document.getElementById('sku-search-results');
    if (!searchInput) {
        resultsDiv.innerHTML = '<div style="color: #64748B; font-size: 11px; padding: 8px;">请输入 SKU 进行搜索</div>';
        return;
    }
    const results = [];
    server_data_cache.forEach(node => {
        if (node.slices && node.slices.length > 0) {
            node.slices.forEach(slice => {
                if (slice.items) {
                    slice.items.forEach(item => {
                        if (item.sku && item.sku.toLowerCase().includes(searchInput)) {
                            results.push({ loc: node.loc, zone: node.zone, sku: item.sku, qty: item.qty, brand: slice.brand });
                        }
                    });
                }
            });
        }
    });
    if (results.length === 0) {
        resultsDiv.innerHTML = `<div style="color: #64748B; font-size: 11px; padding: 8px;">未找到包含 "${searchInput}" 的 SKU</div>`;
    } else {
        let html = `<div style="font-size: 11px; font-weight: bold; margin-bottom: 6px; color: #0F172A;">找到 ${results.length} 个结果:</div>`;
        results.forEach(r => {
            const color = GLOBAL_COLOR_POOL[r.brand] || '#CBD5E1';
            html += `<div style="display: flex; align-items: center; gap: 6px; padding: 6px; background: #F1F5F9; border-radius: 4px; margin-bottom: 4px;">
                <div style="width: 12px; height: 12px; border-radius: 2px; background: ${color}; flex-shrink: 0;"></div>
                <div style="flex: 1; min-width: 0;">
                    <div style="font-weight: bold; font-size: 11px;">${r.sku}</div>
                    <div style="font-size: 10px; color: #64748B;">${r.loc} | ${r.brand}</div>
                </div>
                <div style="font-size: 11px; font-weight: bold; color: #0F172A;">${r.qty}</div>
            </div>`;
        });
        resultsDiv.innerHTML = html;
    }
}

async function loadCloudConfig() {
    if (!CONFIG_API_URL || CONFIG_API_URL === 'null') return;
    try {
        const res = await fetch(CONFIG_API_URL + '?t=' + Date.now());
        const data = await res.json();
        if (data.runtime_config) runtime_config = data.runtime_config;
        if (data.cell_override_db) cell_override_db = data.cell_override_db;
        if (data.actual_brand_colors) { 
            actual_brand_colors = data.actual_brand_colors; 
            Object.assign(GLOBAL_COLOR_POOL, actual_brand_colors); 
        }
        console.log("✅ 云端配置(API)实时加载成功！");
        renderControlPanel();
        applyAllDBCacheToCanvas();
        lockAllEditBtns(); 
    } catch (e) { console.warn("⚠️ 加载云端配置失败: ", e); }
}

function syncConfigToCloud() {
    localStorage.setItem("warehouse_twin_actual_colors_2026", JSON.stringify(actual_brand_colors));
    if (!CONFIG_API_URL || CONFIG_API_URL === 'null') return;
    fetch(CONFIG_API_URL, { method: 'POST', body: JSON.stringify({ key: 'runtime_config', value: runtime_config }), headers: { 'Content-Type': 'text/plain;charset=utf-8' } }).catch(err => console.error('Sync failed:', err));
    fetch(CONFIG_API_URL, { method: 'POST', body: JSON.stringify({ key: 'cell_override_db', value: cell_override_db }), headers: { 'Content-Type': 'text/plain;charset=utf-8' } }).catch(err => console.error('Sync failed:', err));
    fetch(CONFIG_API_URL, { method: 'POST', body: JSON.stringify({ key: 'actual_brand_colors', value: actual_brand_colors }), headers: { 'Content-Type': 'text/plain;charset=utf-8' } }).catch(err => console.error('Sync failed:', err));
}

setInterval(function() { let now = new Date(); let minute = now.getMinutes(); if (minute % 30 == 0 && now.getSeconds() < 10) { window.location.href = window.location.pathname + '?t=' + Date.now(); } }, 5000);
function forceRefreshData() { if(confirm(t('confirmRefresh'))) { window.location.href = window.location.pathname + '?t=' + Date.now(); } }

function switchGlobalView(viewMode) {
    GLOBAL_CURRENT_VIEW = viewMode;
    document.getElementById("view-plan-btn").className = "switch-btn " + (viewMode==='PLAN'?' active':'');
    document.getElementById("view-actual-btn").className = "switch-btn " + (viewMode==='ACTUAL'?' active-actual':'');
    if(viewMode === 'PLAN') { 
        document.getElementById("legend-panel-title").innerText = t('planTitle'); 
        document.getElementById("planning-tools-box").style.display = "block"; 
        document.getElementById("add-brand-btn").style.display = "block";
        document.getElementById("sku-search-box").style.display = "none";
    } else { 
        document.getElementById("legend-panel-title").innerText = t('actualTitle'); 
        document.getElementById("planning-tools-box").style.display = "none"; 
        document.getElementById("add-brand-btn").style.display = "none";
        document.getElementById("sku-search-box").style.display = "block";
    }
    applyAllDBCacheToCanvas(); renderControlPanel();
}

function renderControlPanel() {
    const listContainer = document.getElementById("legend-list"); 
    listContainer.innerHTML = "";
    if (GLOBAL_CURRENT_VIEW === "PLAN") { 
        runtime_config.forEach(item => appendLegendRow(listContainer, item.label, item.color, item.org_name)); 
    } else {
        let sortedBrands = Object.keys(actual_total_stats).sort((a, b) => actual_total_stats[b] - actual_total_stats[a]);
        sortedBrands.forEach(bName => {
            let qty = actual_total_stats[bName];
            let percent = actual_total_qty > 0 ? ((qty / actual_total_qty) * 100).toFixed(1) : '0.0';
            let color = GLOBAL_COLOR_POOL[bName] || '#CBD5E1';
            appendLegendRow(listContainer, bName, color, bName, qty, percent + '%');
        });
        let totalRow = document.createElement("div");
        totalRow.style.cssText = "display:flex; justify-content:space-between; padding:6px 8px; font-size:12px; font-weight:bold; color:#0F172A; border-top:2px solid #CBD5E1; margin-top:6px; background:#F1F5F9; border-radius:4px;";
        totalRow.innerHTML = `<span>${t('totalInventory')}</span><span>${actual_total_qty.toLocaleString()}</span>`;
        listContainer.appendChild(totalRow);
        let occupancyRow = document.createElement("div");
        occupancyRow.style.cssText = "display:flex; justify-content:space-between; padding:6px 8px; font-size:12px; font-weight:bold; color:#0F172A; border-top:1px solid #CBD5E1; margin-top:4px; background:#F1F5F9; border-radius:4px;";
        occupancyRow.innerHTML = `<span>${t('occupancyRate')}</span><span>${occupied_locations} / ${total_locations} (${occupancy_rate}%)</span>`;
        listContainer.appendChild(occupancyRow);
    }
}

function appendLegendRow(container, name, color, orgName, qty, percent) {
    const row = document.createElement("div"); row.style.cssText = "display:flex; align-items:center; gap:6px; background:#F8FAFC; padding:5px 8px; border-radius:6px; border:1px solid #E2E8F0;";
    const colorBox = document.createElement("div"); colorBox.style.cssText = `width:22px; height:20px; border-radius:4px; border:1px solid #CBD5E1; background:${color}; cursor:pointer; flex-shrink:0;`;
    colorBox.onclick = function(e) {
        e.stopPropagation();
        if (!isUnlocked) { alert(t('unlockAlert')); return; }
        const input = document.createElement('input'); input.type = 'color'; input.value = color; input.style.opacity='0';
        input.onchange = function() { updateBrandColor(orgName || name, input.value); };
        colorBox.appendChild(input); input.click();
    };
    const label = document.createElement("span"); label.style.cssText = "font-size:11px; font-weight:bold; flex:1; min-width:0; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;"; label.innerText = name; label.title = name;
    let statsSpan = document.createElement("span"); statsSpan.style.cssText = "font-size:10px; color:#64748B; white-space:nowrap; flex-shrink:0;";
    if (qty !== undefined && percent !== undefined) statsSpan.innerText = `${qty.toLocaleString()} (${percent})`;

    let editBtn = document.createElement("button"); editBtn.innerText = "✏️"; editBtn.className = "lockable"; editBtn.style.cssText = "background:none; border:none; cursor:pointer; flex-shrink:0; font-size:14px;"; editBtn.onclick = function(e) { e.stopPropagation(); editBrand(orgName || name); };
    let delBtn = document.createElement("button"); delBtn.innerText = "🗑️"; delBtn.className = "lockable"; delBtn.style.cssText = "background:none; border:none; cursor:pointer; flex-shrink:0; font-size:14px;"; delBtn.onclick = function(e) { e.stopPropagation(); if(confirm(`${t('deleteConfirm')} "${name}"?`)) deleteBrand(orgName || name); };
    let resetBtn = document.createElement("button"); resetBtn.innerText = "🔄"; resetBtn.className = "lockable"; resetBtn.style.cssText = "background:none; border:none; cursor:pointer; flex-shrink:0; font-size:14px;"; resetBtn.onclick = function(e) { e.stopPropagation(); if(confirm(`${t('restoreConfirm')} "${name}"?`)) resetBrandLocations(orgName || name); };

    row.appendChild(colorBox); row.appendChild(label);
    if (qty !== undefined) row.appendChild(statsSpan);
    if (GLOBAL_CURRENT_VIEW === 'PLAN') { row.appendChild(editBtn); row.appendChild(delBtn); row.appendChild(resetBtn); }
    container.appendChild(row);
}

function deleteBrand(brandName) { runtime_config = runtime_config.filter(c => c.org_name !== brandName); applyAllDBCacheToCanvas(); renderControlPanel(); syncConfigToCloud(); }
function editBrand(brand) {
    let locs = [];
    server_data_cache.forEach(node => { let currentBrand = node.native_brand; if (cell_override_db[node.loc]) currentBrand = cell_override_db[node.loc].org_name; if (currentBrand === brand) locs.push(node.loc); });
    document.getElementById('target-loc').value = formatLocsToRange(locs); document.getElementById('new-brand').value = brand;
    let conf = runtime_config.find(c => c.org_name === brand || c.label === brand);
    document.getElementById('new-color').value = (conf && conf.color !== 'transparent') ? conf.color : "#FFFFFF";
    document.getElementById('planning-tools-box').scrollIntoView({behavior: 'smooth'}); 
}
function resetBrandLocations(brand) { server_data_cache.forEach(node => { if (node.native_brand === brand) delete cell_override_db[node.loc]; }); applyAllDBCacheToCanvas(); syncConfigToCloud(); }
function formatLocsToRange(locs) { 
    if (!locs || locs.length === 0) return ""; 
    let parsed = locs.map(l => { let m = l.match(/^([A-Z]+)(\d+)-(\d+)$/); return m ? { raw: l, z: m[1], c: parseInt(m[2]), l: parseInt(m[3]) } : { raw: l, z: l, c: 0, l: 0 }; }); 
    parsed.sort((a, b) => a.z.localeCompare(b.z) || a.c - b.c || a.l - b.l); 
    let ranges = [], i = 0; 
    while (i < parsed.length) { let start = parsed[i], end = parsed[i]; while (i + 1 < parsed.length && parsed[i+1].z === start.z && parsed[i+1].c === start.c && parsed[i+1].l === end.l + 1) { i++; end = parsed[i]; } ranges.push(start.raw === end.raw ? start.raw : `${start.raw}~${end.raw}`); i++; } 
    return ranges.join(", "); 
}
function updateBrandColor(brand, newColor) { 
    let conf = runtime_config.find(c => c.org_name === brand); 
    if (conf) conf.color = newColor; else actual_brand_colors[brand] = newColor;
    GLOBAL_COLOR_POOL[brand] = newColor; applyAllDBCacheToCanvas(); renderControlPanel(); syncConfigToCloud(); 
}
function applyAllDBCacheToCanvas() { 
    var gd = document.getElementsByClassName('plotly-graph-div')[0]; if(!gd) return; 
    let originalTraces = gd.data, basePlatformAndScatters = [], shelfCubesMap = {}; 
    for(let i=0; i<originalTraces.length; i++) { let t = originalTraces[i]; if(t.name === '_SHELF_CUBE_' || t.name === '_DYNAMIC_SLICE_') { if(t.customdata) shelfCubesMap[t.customdata[0]] = t; } else basePlatformAndScatters.push(t); } 
    let finalDynamicTraces = [...basePlatformAndScatters]; 
    server_data_cache.forEach(node => { 
        let locID = node.loc, templateCube = shelfCubesMap[locID]; if(!templateCube) return; 
        if (GLOBAL_CURRENT_VIEW === "PLAN") { 
            let targetColor = node.native_color, targetLabel = node.native_brand; 
            if(cell_override_db[locID]) { targetColor = cell_override_db[locID].color; targetLabel = cell_override_db[locID].label; } 
            else { let parentConf = runtime_config.find(c => c.org_name === node.native_brand); if(parentConf) { targetColor = parentConf.color || GLOBAL_COLOR_POOL[parentConf.org_name]; targetLabel = parentConf.label; } else { targetColor = GLOBAL_COLOR_POOL[node.native_brand] || node.native_color; targetLabel = node.native_brand; } } 
            templateCube.color = targetColor; templateCube.name = '_SHELF_CUBE_'; templateCube.z = node.orig_z; templateCube.text = Array(8).fill(`<b>${locID}</b><br>${targetLabel}`); finalDynamicTraces.push(templateCube); 
        } else { 
            let slices = node.slices || []; var origZ = node.orig_z; var minZ = Math.min(...origZ), maxZ = Math.max(...origZ), fullHeight = maxZ - minZ; 
            for(let s=0; s<slices.length; s++) { 
                let currentSlice = slices[s], segmentHeight = fullHeight / slices.length; let sliceMinZ = minZ + (s * segmentHeight), sliceMaxZ = sliceMinZ + segmentHeight; 
                let currentZArray = [...origZ]; for(let v=0; v<8; v++) { if(origZ[v] === minZ) currentZArray[v] = sliceMinZ; else currentZArray[v] = sliceMaxZ; } 
                let sliceColor = GLOBAL_COLOR_POOL[currentSlice.brand] || currentSlice.color; let hoverHTML = `<b>${locID}</b><br>${currentSlice.brand}<br>`; 
                currentSlice.items.forEach(it => { hoverHTML += `${it.sku}: ${it.qty}<br>`; }); 
                finalDynamicTraces.push({ type: 'mesh3d', x: templateCube.x, y: templateCube.y, z: currentZArray, i: templateCube.i, j: templateCube.j, k: templateCube.k, color: sliceColor, customdata: Array(8).fill(locID), text: Array(8).fill(hoverHTML), name: '_DYNAMIC_SLICE_', hoverinfo: 'text', showlegend: false }); 
            } 
        } 
    }); gd.data = finalDynamicTraces; Plotly.redraw(gd); 
}
function parseSinglePattern(pat, locID) { 
    pat = pat.trim().toUpperCase(); if (!pat) return false; 
    if (pat.includes('~')) { let parts = pat.split('~'); if (parts.length === 2) { let mS = parts[0].match(/^([A-Z]+)(\d+)-(\d+)$/), mE = parts[1].match(/^([A-Z]+)(\d+)-(\d+)$/), mL = locID.match(/^([A-Z]+)(\d+)-(\d+)$/); if (mS && mE && mL && mL[1] === mS[1]) return (parseInt(mL[2]) >= Math.min(parseInt(mS[2]), parseInt(mE[2])) && parseInt(mL[2]) <= Math.max(parseInt(mS[2]), parseInt(mE[2])) && parseInt(mL[3]) >= Math.min(parseInt(mS[3]), parseInt(mE[3])) && parseInt(mL[3]) <= Math.max(parseInt(mS[3]), parseInt(mE[3]))); } } 
    return locID === pat; 
}
function applyLocationChange() { 
    let raw = document.getElementById('target-loc').value.trim(), brand = document.getElementById('new-brand').value.trim(), color = document.getElementById('new-color').value; 
    if(!raw || !brand) return; let pats = raw.split(','); let exist = runtime_config.some(c => c.org_name === brand); let fOrg = exist ? runtime_config.find(c => c.org_name === brand).org_name : brand; 
    if (!exist) runtime_config.push({ org_name: brand, color: color, label: brand }); else { let conf = runtime_config.find(c => c.org_name === brand); if(conf) conf.color = color; }
    GLOBAL_COLOR_POOL[fOrg] = color; server_data_cache.forEach(node => { if(pats.some(p => parseSinglePattern(p, node.loc))) cell_override_db[node.loc] = { org_name: fOrg, label: brand, color: color }; }); 
    document.getElementById('target-loc').value = ""; document.getElementById('new-brand').value = ""; applyAllDBCacheToCanvas(); renderControlPanel(); syncConfigToCloud(); 
}
function addNewBrand() {
    let brandName = prompt(t('promptBrandName')); if (!brandName || !brandName.trim()) return; brandName = brandName.trim();
    if (runtime_config.some(c => c.org_name === brandName)) { alert(t('promptBrandExists')); return; }
    let colorHex = prompt(t('promptBrandColor'), "#3B82F6"); if (colorHex === null) return; if (!colorHex.trim()) colorHex = "#3B82F6";
    runtime_config.push({ org_name: brandName, color: colorHex, label: brandName }); GLOBAL_COLOR_POOL[brandName] = colorHex; applyAllDBCacheToCanvas(); renderControlPanel(); syncConfigToCloud();
}
function resetToDefault() { 
    if(confirm(t('confirmReset'))) { 
        if (CONFIG_API_URL) { fetch(CONFIG_API_URL, { method: 'POST', body: JSON.stringify({ key: 'runtime_config', value: [] }), headers: { 'Content-Type': 'text/plain' } }); fetch(CONFIG_API_URL, { method: 'POST', body: JSON.stringify({ key: 'cell_override_db', value: {} }), headers: { 'Content-Type': 'text/plain' } }); fetch(CONFIG_API_URL, { method: 'POST', body: JSON.stringify({ key: 'actual_brand_colors', value: {} }), headers: { 'Content-Type': 'text/plain' } }); }
        localStorage.removeItem("warehouse_twin_master_2026"); localStorage.removeItem("warehouse_twin_cell_overrides_2026"); localStorage.removeItem("warehouse_twin_actual_colors_2026"); window.location.reload(); 
    } 
}
function toggleNav() { document.getElementById('super-legend-panel').classList.toggle('nav-open'); }
function setMode(mode, btn) {
    var gd = document.getElementsByClassName('plotly-graph-div')[0]; if (!gd) return;
    var currentCamera = gd.layout.scene.camera || {eye: {x: -0.8, y: -0.8, z: 3.5}, center: {x: 0, y: 0, z: 0}, up: {x: 0, y: 0, z: 1}};
    if (btn.classList.contains('active')) { mode = 'turntable'; document.getElementById('btn-rotate').classList.add('active'); document.getElementById('btn-pan').classList.remove('active'); } else { document.querySelectorAll('.ctrl-btn').forEach(b => { if(b.id.startsWith('btn-')) b.classList.remove('active'); }); btn.classList.add('active'); }
    var eye = currentCamera.eye; var center = currentCamera.center || {x: 0, y: 0, z: 0}; var slowFactor = (mode === 'pan') ? 2.0 : 0.5; 
    var newEye = { x: center.x + (eye.x - center.x) * slowFactor, y: center.y + (eye.y - center.y) * slowFactor, z: center.z + (eye.z - center.z) * slowFactor };
    var update = {'scene.dragmode': mode}; update['scene.camera.eye'] = newEye; update['scene.camera.center'] = center; update['scene.camera.up'] = currentCamera.up || {x: 0, y: 0, z: 1}; Plotly.relayout(gd, update);
}
let currentScale = 1.0; 
function zoomCamera(factor) { if (factor < 1) { currentScale = currentScale * 1.2; } else { currentScale = currentScale * 0.8; } if (currentScale > 3.0) currentScale = 3.0; if (currentScale < 0.5) currentScale = 0.5; var plotContainer = document.querySelector('.plotly-graph-div'); if (plotContainer) { plotContainer.style.transform = `scale(${currentScale})`; plotContainer.style.transformOrigin = 'center center'; } }
function resetCamera() { currentScale = 1.0; var plotContainer = document.querySelector('.plotly-graph-div'); if (plotContainer) { plotContainer.style.transform = 'scale(1)'; } var gd = document.getElementsByClassName('plotly-graph-div')[0]; if (gd) { Plotly.relayout(gd, { 'scene.camera.eye': {x: -0.8, y: -0.8, z: 3.5}, 'scene.camera.center': {x: 0, y: 0, z: 0} }); } }
const TARGET_HASH = "f0a36b9da192dc4732c232774766160f204bfe18be84c0a0dafce7040334b29f"; let isUnlocked = false;
function showPwdModal() { if (isUnlocked) { isUnlocked = false; lockAllEditBtns(); document.querySelector('#control-panel button:last-child').innerText = "🔒"; } else { document.getElementById('pwd-modal').style.display = 'flex'; document.getElementById('pwd-input').value = ''; document.getElementById('pwd-input').focus(); } }
function closePwdModal() { document.getElementById('pwd-modal').style.display = 'none'; }
async function verifyPwd() { const pwd = document.getElementById('pwd-input').value; if (!pwd) return; const hashBuffer = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(pwd)); const hashHex = Array.from(new Uint8Array(hashBuffer)).map(b => b.toString(16).padStart(2, '0')).join(''); if (hashHex === TARGET_HASH) { isUnlocked = true; closePwdModal(); unlockAllEditBtns(); document.querySelector('#control-panel button:last-child').innerText = "🔓"; } else { alert(t('wrongPwd')); } }
function lockAllEditBtns() { document.querySelectorAll('.lockable').forEach(btn => btn.classList.add('locked')); }
function unlockAllEditBtns() { document.querySelectorAll('.lockable').forEach(btn => btn.classList.remove('locked')); }

if (window.innerWidth <= 768) { document.getElementById('super-legend-panel').classList.remove('nav-open'); }
var checkPlotly = setInterval(function(){ 
    var gd = document.getElementsByClassName('plotly-graph-div')[0]; 
    if(gd && gd._fullLayout) { 
        clearInterval(checkPlotly); 
        applyLanguage(); 
        renderControlPanel(); 
        applyAllDBCacheToCanvas(); 
        lockAllEditBtns(); 
        loadCloudConfig(); 
    }
}, 400);
</script>
'''

    replacements = {
        "SERVER_CONFIG_INJECT_PLACEHOLDER": js_config_string,
        "SERVER_OVERRIDES_INJECT_PLACEHOLDER": js_overrides_string,
        "ACTUAL_COLORS_INJECT_PLACEHOLDER": js_actual_colors_string,
        "ACTUAL_TOTAL_STATS_PLACEHOLDER": js_actual_total_stats,
        "ACTUAL_TOTAL_QTY_PLACEHOLDER": js_actual_total_qty,
        "TOTAL_LOCATIONS_PLACEHOLDER": js_total_locations,
        "OCCUPIED_LOCATIONS_PLACEHOLDER": js_occupied_locations,
        "OCCUPANCY_RATE_PLACEHOLDER": js_occupancy_rate,
        "CONFIG_API_URL_PLACEHOLDER": js_api_url_string,
        "SERVER_DATA_INJECT_PLACEHOLDER": js_array_string,
        "SERVER_COLORS_INJECT_PLACEHOLDER": js_global_colors_string,
        "DATA_TIMESTAMP_PLACEHOLDER": data_timestamp
    }
    for key, value in replacements.items():
        interactive_control_script = interactive_control_script.replace(key, str(value))

    final_html = html_content.replace("</body>", interactive_control_script + "</body>")
    with open(OUTPUT_HTML, "w", encoding="utf-8") as f: f.write(final_html)

if __name__ == "__main__":
    is_github_actions = os.environ.get('GITHUB_ACTIONS') == 'true'
    if is_github_actions:
        print("="*50)
        print(" GitHub Actions 模式：单次运行 ")
        print("="*50)
        try: 
            generate_html()
            print("✅ 沙盘更新完成！ ")
        except Exception as e: 
            print(f" 生成失败: {e} ")
            import traceback
            traceback.print_exc()
            exit(1)
    else:
        print("="*50)
        print(" 仓库沙盘双核系统已启动，进入后台监控模式... ")
        print("💡 程序将每 30 分钟自动抓取最新数据。 ")
        print(" 请保持此窗口运行。按 Ctrl+C 可安全退出程序。 ")
        print("="*50)
        last_run_timestamp = 0
        REFRESH_INTERVAL_SECONDS = 30 * 60
        try:
            while True:
                now = datetime.datetime.now()
                current_timestamp = time.time()
                if last_run_timestamp == 0 or not os.path.exists(OUTPUT_HTML) or (current_timestamp - last_run_timestamp) >= REFRESH_INTERVAL_SECONDS:
                    print(f"\n🔄 [{now.strftime('%Y-%m-%d %H:%M:%S')}] 触发刷新机制，正在生成最新沙盘... ")
                    try: 
                        generate_html()
                        last_run_timestamp = current_timestamp
                        print("✅ 沙盘更新完成！等待下一次 30 分钟周期... ")
                    except Exception as e: 
                        print(f" 生成失败: {e} ")
                        import traceback
                        traceback.print_exc()
                time.sleep(60)
        except KeyboardInterrupt: 
            print("\n\n👋 收到退出信号，仓库沙盘后台监控已安全停止。 ")