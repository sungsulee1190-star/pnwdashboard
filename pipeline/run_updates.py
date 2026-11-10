import os
import sys
import glob
import re
import datetime
import win32com.client
import pandas as pd
import subprocess

# Configure stdout for utf-8 to avoid console encoding issues
sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

DRY_RUN = "--dry-run" in sys.argv
SKIP_GIT = "--no-git" in sys.argv

import json

# 코드 위치: <저장소>/pipeline/run_updates.py. 엑셀은 paths.local.json의 source_dir에서만 읽는다.
PIPELINE_DIR = os.path.dirname(os.path.abspath(__file__))
DASHBOARD_DIR = os.path.dirname(PIPELINE_DIR)
WORKSPACE_DIR = os.path.dirname(DASHBOARD_DIR)  # 저장소 바깥 작업 폴더(Dashboard_Backup 위치)
PATHS_FILE = os.path.join(PIPELINE_DIR, "paths.local.json")

if not os.path.exists(PATHS_FILE):
    sys.exit(f"[중단] 경로 설정 파일이 없습니다: {PATHS_FILE}\n"
             f"       paths.example.json을 paths.local.json으로 복사하고 source_dir를 고치세요.")
with open(PATHS_FILE, encoding="utf-8") as f:
    PATHS = json.load(f)

SOURCE_DIR = PATHS["source_dir"]
MASTER_WEEKLY_PATH = os.path.join(SOURCE_DIR, PATHS["master_weekly"])
MASTER_CN_PATH = os.path.join(SOURCE_DIR, PATHS["master_cn"])
MASTER_KR_PATH = os.path.join(SOURCE_DIR, PATHS["master_kr"])
MASTER_PATHS = [MASTER_WEEKLY_PATH, MASTER_CN_PATH, MASTER_KR_PATH]

# 원본(raw) 입력 폴더
RAW_CN_DIR = os.path.join(SOURCE_DIR, PATHS["raw_cn_dir"])
RAW_KR_DIR = os.path.join(SOURCE_DIR, PATHS["raw_kr_dir"])
DOMESTIC_DIR = os.path.join(SOURCE_DIR, PATHS["domestic_dir"])

DATA_JS_PATH = os.path.join(DASHBOARD_DIR, "data.js")

# 다른 폴더의 사본으로 대체하지 않는다. 하나라도 없으면 이름을 밝히고 멈춘다.
_missing = [p for p in [SOURCE_DIR, *MASTER_PATHS, RAW_CN_DIR, RAW_KR_DIR, DOMESTIC_DIR] if not os.path.exists(p)]
if _missing:
    sys.exit("[중단] 엑셀 폴더에 다음 경로가 없습니다 (paths.local.json 확인):\n  - " + "\n  - ".join(_missing))

def open_wb_safe(excel, path, read_only=False):
    """Opens an Excel workbook safely, reusing the already-open workbook if open in the user's Excel."""
    target_name = os.path.basename(path).lower()
    try:
        for wb in excel.Workbooks:
            if wb.Name.lower() == target_name:
                return wb, True
    except Exception:
        pass
    wb = excel.Workbooks.Open(path, UpdateLinks=0, ReadOnly=read_only)
    return wb, False

def close_wb_safe(wb, was_open):
    """Closes workbook only if it was not already open by the user."""
    if wb is not None and not was_open:
        try:
            wb.Close(False)
        except Exception:
            pass

# ─────────────────────────────────────────────
# 2. File Utilities
# ─────────────────────────────────────────────
def robust_find_file(folder, keywords, extension):
    """Finds files matching multiple keywords in folder and all subfolders. Latest period in the file name wins; modification time breaks ties."""
    if not os.path.exists(folder):
        print(f"  [Error] Folder does not exist: {folder}")
        return None
    
    files = []
    for root, dirs, fnames in os.walk(folder):
        # Skip hidden/temp directories
        dirs[:] = [d for d in dirs if not d.startswith('.')]
        for fn in fnames:
            if fn.startswith("~$"):
                continue
            if fn.endswith(extension):
                match_all = all(kw in fn for kw in keywords)
                if match_all:
                    files.append(os.path.join(root, fn))
                
    if not files:
        return None
    def rank(path):
        year, month = parse_date_from_filename(os.path.basename(path))
        return (year or 0, month or 0, os.path.getmtime(path))
    return max(files, key=rank)

def parse_date_from_filename(filename):
    """Parses year and month/day from filename using regular expressions."""
    # Matches patterns like 2026年5월 or 2026.4월 or 2026.12월 or 2026년05월
    match = re.search(r"(\d{4})[-.\s]*(?:년|年)?\s*(\d{1,2})\s*(?:월|月)?", filename)
    if match:
        year = int(match.group(1))
        month = int(match.group(2))
        return year, month
    return None, None

def parse_excel_date(cell_val):
    """Parses dates from Excel cells robustly, supporting datetime, string, and numeric float representations."""
    if not cell_val:
        return None, None
    # Check for datetime or pywintypes.datetime (has year/month attributes)
    if hasattr(cell_val, 'year') and hasattr(cell_val, 'month'):
        return cell_val.year, cell_val.month
    
    val_str = str(cell_val).strip()
    # Pattern YYYY-MM
    match = re.match(r"(\d{4})[-/](\d{1,2})", val_str)
    if match:
        return int(match.group(1)), int(match.group(2))
    # Pattern YYYY년 MM월
    match = re.match(r"(\d{4})년\s*(\d{1,2})월", val_str)
    if match:
        return int(match.group(1)), int(match.group(2))
    # Pattern YY년 MM월
    match = re.match(r"(\d{2})년\s*(\d{1,2})월", val_str)
    if match:
        return 2000 + int(match.group(1)), int(match.group(2))
    return None, None

# ─────────────────────────────────────────────
# 3. Excel Master Updates via Excel COM
# ─────────────────────────────────────────────
def update_weekly_price_master(excel, raw_weekly_path):
    """Updates the weekly price master workbook with the latest weekly price column."""
    master_path = MASTER_WEEKLY_PATH
    print(f"\n[Part 1] 중국 주간 가격 마스터 업데이트")
    print(f"  - 마스터 파일: {os.path.basename(master_path)}")
    print(f"  - 원본 주간 파일: {os.path.basename(raw_weekly_path)}")
    
    # Open workbooks
    wb_master, was_open_m = open_wb_safe(excel, master_path, read_only=False)
    try:
        excel.Calculation = -4135 # xlCalculationManual = -4135
    except Exception as e:
        print(f"  [Warning] Calculation mode change failed: {e}")
    wb_raw, was_open_r = open_wb_safe(excel, raw_weekly_path, read_only=True)
    
    try:
        # 1. Determine latest date column in Raw sheet
        ws_raw_coated = wb_raw.Sheets("铜版纸")
        used_raw_coated = ws_raw_coated.UsedRange
        c_raw_count = used_raw_coated.Columns.Count
        
        # Read headers from Row 2 to find the latest date
        latest_date_val = None
        latest_col_idx = None
        for col in range(c_raw_count, 5, -1): # Dates start after column 5
            val = ws_raw_coated.Cells(2, col).Value
            if val and str(val).strip() and val != "/":
                latest_date_val = val
                latest_col_idx = col
                break
                
        if not latest_date_val or not latest_col_idx:
            print("  [Error] 원본 파일에서 유효한 주간 날짜 컬럼을 찾지 못했습니다.")
            return
            
        # Format date header as YYYY-MM-DD string
        if hasattr(latest_date_val, 'strftime'):
            date_str = latest_date_val.strftime("%Y-%m-%d")
        else:
            date_str = str(latest_date_val).strip().split(" ")[0]
            
        print(f"  - 최신 주간 날짜 확인: {date_str} (컬럼 인덱스 {latest_col_idx})")
        
        # Build Raw Price Dictionaries for quick lookup
        # Map: code -> price
        raw_prices = {}
        for ws_name in ["铜版纸", "双胶纸"]:
            ws_raw = wb_raw.Sheets(ws_name)
            used_r = ws_raw.UsedRange
            r_cnt = used_r.Rows.Count
            for r in range(3, r_cnt + 1):
                code = ws_raw.Cells(r, 1).Value
                price = ws_raw.Cells(r, latest_col_idx).Value
                if code:
                    raw_prices[str(code).strip()] = price
                    
        # Map: (region, brand, gsm) -> price for 双胶纸 sheet lookup by name
        raw_names_prices = {}
        ws_raw_sz = wb_raw.Sheets("双胶纸")
        used_sz = ws_raw_sz.UsedRange
        for r in range(3, used_sz.Rows.Count + 1):
            region = ws_raw_sz.Cells(r, 3).Value
            brand = ws_raw_sz.Cells(r, 4).Value
            gsm = ws_raw_sz.Cells(r, 5).Value
            price = ws_raw_sz.Cells(r, latest_col_idx).Value
            if region and brand and gsm:
                key = (str(region).strip(), str(brand).strip(), str(gsm).strip())
                raw_names_prices[key] = price
                
        # 2. Update Master Sheet: ' Raw_铜版纸(C2S), 백상지(双胶纸)'
        ws_m_raw = wb_master.Sheets(' Raw_铜版纸(C2S), 백상지(双胶纸)')
        used_m_raw = ws_m_raw.UsedRange
        m_r_count = used_m_raw.Rows.Count
        m_c_count = used_m_raw.Columns.Count
        
        # Check if column already exists
        exists_col = None
        for col in range(11, m_c_count + 1):
            h_val = ws_m_raw.Cells(1, col).Value
            h_str = h_val.strftime("%Y-%m-%d") if hasattr(h_val, 'strftime') else str(h_val).strip().split(" ")[0]
            if h_str == date_str:
                exists_col = col
                break
                
        if exists_col:
            print(f"  - Sheet ' Raw_...': {date_str} 컬럼이 이미 존재합니다 (컬럼 {exists_col}). 덮어씁니다.")
            target_col = exists_col
        else:
            print(f"  - Sheet ' Raw_...': {date_str} 새 컬럼을 추가합니다 (컬럼 {m_c_count + 1}).")
            target_col = m_c_count + 1
            ws_m_raw.Cells(1, target_col).Value = date_str
            # Copy format from previous column
            ws_m_raw.Columns(target_col - 1).Copy()
            ws_m_raw.Columns(target_col).PasteSpecial(Paste=-4122) # xlPasteFormats = -4122
            excel.CutCopyMode = False
            
        # Update values
        for r in range(2, m_r_count + 1):
            code = ws_m_raw.Cells(r, 6).Value # Column 6 is 编码
            if code:
                code_str = str(code).strip()
                if code_str in raw_prices:
                    ws_m_raw.Cells(r, target_col).Value = raw_prices[code_str]
                    
        # 3. Update Master Sheet: '双胶纸'
        ws_m_sz = wb_master.Sheets('双胶纸')
        used_m_sz = ws_m_sz.UsedRange
        m_sz_r_count = used_m_sz.Rows.Count
        m_sz_c_count = used_m_sz.Columns.Count
        
        # Check if column already exists
        exists_sz_col = None
        for col in range(6, m_sz_c_count + 1):
            h_val = ws_m_sz.Cells(2, col).Value # Header is in Row 2!
            h_str = h_val.strftime("%Y-%m-%d") if hasattr(h_val, 'strftime') else str(h_val).strip().split(" ")[0]
            if h_str == date_str:
                exists_sz_col = col
                break
                
        if exists_sz_col:
            print(f"  - Sheet '双胶纸': {date_str} 컬럼이 이미 존재합니다 (컬럼 {exists_sz_col}). 덮어씁니다.")
            target_sz_col = exists_sz_col
        else:
            print(f"  - Sheet '双胶纸': {date_str} 새 컬럼을 추가합니다 (컬럼 {m_sz_c_count + 1}).")
            target_sz_col = m_sz_c_count + 1
            ws_m_sz.Cells(2, target_sz_col).Value = date_str
            # Copy format from previous column
            ws_m_sz.Columns(target_sz_col - 1).Copy()
            ws_m_sz.Columns(target_sz_col).PasteSpecial(Paste=-4122)
            excel.CutCopyMode = False
            
        # Update values
        for r in range(3, m_sz_r_count + 1):
            region = ws_m_sz.Cells(r, 3).Value # Column C is 地区
            brand = ws_m_sz.Cells(r, 4).Value  # Column D is 品牌
            gsm = ws_m_sz.Cells(r, 5).Value    # Column E is 克重
            if region and brand and gsm:
                key = (str(region).strip(), str(brand).strip(), str(gsm).strip())
                if key in raw_names_prices:
                    ws_m_sz.Cells(r, target_sz_col).Value = raw_names_prices[key]
                    
        if not DRY_RUN:
            wb_master.Save()
            print("  ✅ 주간 가격 마스터 업데이트 완료!")
        else:
            print("  [Dry-run] 주간 가격 마스터 저장을 건너뜁니다.")
    finally:
        close_wb_safe(wb_raw, was_open_r)
        close_wb_safe(wb_master, was_open_m)

def update_china_monthly_master(excel, raw_monthly_path):
    """Updates the Chinese monthly master workbook (prices, indicators, and country export data)."""
    master_path = MASTER_CN_PATH
    print(f"\n[Part 2] 중국 월간 실적 마스터 업데이트")
    print(f"  - 마스터 파일: {os.path.basename(master_path)}")
    print(f"  - 원본 월보 파일: {os.path.basename(raw_monthly_path)}")
    
    # Open workbooks
    wb_master, was_open_m = open_wb_safe(excel, master_path, read_only=False)
    try:
        excel.Calculation = -4135 # xlCalculationManual = -4135
    except Exception as e:
        print(f"  [Warning] Calculation mode change failed: {e}")
    wb_raw, was_open_r = open_wb_safe(excel, raw_monthly_path, read_only=True)
    
    try:
        # Determine raw sheets
        raw_sheet_names = [s.Name for s in wb_raw.Sheets]
        
        # ─────────────────────────────────────────────────────────────
        # 2A. UMPaper 가격 Index 시트 업데이트
        # ─────────────────────────────────────────────────────────────
        ws_m_index = wb_master.Sheets('UMPaper文化纸价格指数 (가격 Index)')
        ws_r_index = wb_raw.Sheets('文化纸价格指数')
        
        used_m_idx = ws_m_index.UsedRange
        c_m_idx = used_m_idx.Columns.Count
        
        # Find latest month from Raw
        used_r_idx = ws_r_index.UsedRange
        c_r_idx = used_r_idx.Columns.Count
        
        latest_month_val = None
        latest_col_idx = None
        for col in range(c_r_idx, 2, -1):
            val = ws_r_index.Cells(2, col).Value
            if val and str(val).strip():
                latest_month_val = str(val).strip()
                latest_col_idx = col
                break
                
        if not latest_month_val:
            print("  [Error] Raw '文化纸价格指数'에서 최신 날짜 컬럼을 찾지 못했습니다.")
            return
            
        print(f"  - 가격지수 최신 월 확인: {latest_month_val}")
        
        # Check if already exists in master
        exists_idx_col = None
        for col in range(4, c_m_idx + 1):
            h_val = ws_m_index.Cells(2, col).Value
            if h_val and str(h_val).strip() == latest_month_val:
                exists_idx_col = col
                break
                
        if exists_idx_col:
            print(f"  - Sheet '가격 Index': {latest_month_val} 컬럼이 존재합니다. 덮어씁니다.")
            target_idx_col = exists_idx_col
        else:
            print(f"  - Sheet '가격 Index': {latest_month_val} 컬럼을 추가합니다.")
            target_idx_col = c_m_idx + 1
            ws_m_index.Cells(2, target_idx_col).Value = latest_month_val
            # Copy formats
            ws_m_index.Columns(target_idx_col - 1).Copy()
            ws_m_index.Columns(target_idx_col).PasteSpecial(Paste=-4122)
            excel.CutCopyMode = False
            
        # Copy prices (Rows 3-11)
        # Create map of raw product name -> price
        raw_prices_map = {}
        for r in range(3, 12):
            prod = ws_r_index.Cells(r, 2).Value
            val = ws_r_index.Cells(r, latest_col_idx).Value
            if prod:
                raw_prices_map[str(prod).strip()] = val
                
        # Fill master index values
        for r in range(3, 12):
            prod = ws_m_index.Cells(r, 3).Value # Column C is product
            if prod:
                prod_str = str(prod).strip()
                if prod_str in raw_prices_map:
                    ws_m_index.Cells(r, target_idx_col).Value = raw_prices_map[prod_str]
                    
        # Copy formulas for Rows 13, 14, 15
        for r in [13, 14, 15]:
            ws_m_index.Cells(r, target_idx_col).Formula = ws_m_index.Cells(r, target_idx_col - 1).Formula
            
        # ─────────────────────────────────────────────────────────────
        # 2B. 주요지표 시트 업데이트
        # ─────────────────────────────────────────────────────────────
        ws_m_ind = wb_master.Sheets('주요지표')
        ws_r_prod = wb_raw.Sheets('文化纸产量')
        ws_r_inv = wb_raw.Sheets('文化纸库存')
        ws_r_coated_trade = wb_raw.Sheets('铜版纸进出口量')
        ws_r_offset_trade = wb_raw.Sheets('双胶纸进出口量')
        
        used_m_ind = ws_m_ind.UsedRange
        c_m_ind = used_m_ind.Columns.Count
        
        # Get target statistics month from raw sheet '文化纸产量'
        # Production data runs 1 month behind prices
        used_r_prod = ws_r_prod.UsedRange
        c_r_prod = used_r_prod.Columns.Count
        target_ind_month = None
        target_r_col = None
        
        # Find latest month column in Raw (ignoring total/cumulative column)
        for col in range(c_r_prod, 1, -1):
            val = ws_r_prod.Cells(2, col).Value
            if val and str(val).strip() and "累计" not in str(val) and "합계" not in str(val):
                target_ind_month = str(val).strip()
                target_r_col = col
                break
                
        if not target_ind_month:
            print("  [Error] Raw '文化纸产量'에서 최신 지표 대상 월을 찾지 못했습니다.")
            return
            
        # Reformat YYYY/MM header to match master YYYY-MM format
        # e.g., 2026/4 -> 2026-04, 2026/04 -> 2026-04
        date_parts = target_ind_month.split("/")
        if len(date_parts) == 2:
            y = int(date_parts[0])
            m = int(date_parts[1])
            ind_header_str = f"{y}-{m:02d}"
        else:
            ind_header_str = target_ind_month
            
        print(f"  - 실적지표 최신 월 확인: {ind_header_str}")
        
        # Check if already exists in master (in Row 3)
        exists_ind_col = None
        for col in range(3, c_m_ind + 1):
            h_val = ws_m_ind.Cells(3, col).Value
            if h_val and str(h_val).strip() == ind_header_str:
                exists_ind_col = col
                break
                
        if exists_ind_col:
            print(f"  - Sheet '주요지표': {ind_header_str} 컬럼이 존재합니다. 덮어씁니다.")
            target_ind_col = exists_ind_col
        else:
            print(f"  - Sheet '주요지표': {ind_header_str} 컬럼을 추가합니다.")
            target_ind_col = c_m_ind + 1
            ws_m_ind.Cells(3, target_ind_col).Value = ind_header_str
            # Copy format
            ws_m_ind.Columns(target_ind_col - 1).Copy()
            ws_m_ind.Columns(target_ind_col).PasteSpecial(Paste=-4122)
            excel.CutCopyMode = False
            
        # Fill production (Rows 20-25)
        # Row 20: 铜版纸产量, Row 21: Coated YoY, Row 22: 双胶纸产量, Row 23: Double Offset YoY, Row 24: 新闻纸产量, Row 25: News YoY
        ws_m_ind.Cells(20, target_ind_col).Value = ws_r_prod.Cells(3, target_r_col).Value
        ws_m_ind.Cells(21, target_ind_col).Value = ws_r_prod.Cells(4, target_r_col).Value
        ws_m_ind.Cells(22, target_ind_col).Value = ws_r_prod.Cells(5, target_r_col).Value
        ws_m_ind.Cells(23, target_ind_col).Value = ws_r_prod.Cells(6, target_r_col).Value
        ws_m_ind.Cells(24, target_ind_col).Value = ws_r_prod.Cells(7, target_r_col).Value
        ws_m_ind.Cells(25, target_ind_col).Value = ws_r_prod.Cells(8, target_r_col).Value
        
        # Fill inventory (Rows 28-32)
        # Target month in inventory raw sheet is YYYY/MM. Look it up.
        r_inv_col = None
        for col in range(ws_r_inv.UsedRange.Columns.Count, 1, -1):
            val = ws_r_inv.Cells(2, col).Value
            if val and str(val).strip() == target_ind_month:
                r_inv_col = col
                break
                
        if r_inv_col:
            ws_m_ind.Cells(28, target_ind_col).Value = target_ind_month
            ws_m_ind.Cells(29, target_ind_col).Value = ws_r_inv.Cells(3, r_inv_col).Value
            ws_m_ind.Cells(30, target_ind_col).Value = ws_r_inv.Cells(4, r_inv_col).Value
            ws_m_ind.Cells(31, target_ind_col).Value = ws_r_inv.Cells(5, r_inv_col).Value
            ws_m_ind.Cells(32, target_ind_col).Value = ws_r_inv.Cells(6, r_inv_col).Value
        else:
            print(f"  ⚠️  Raw '文化纸库存'에서 {target_ind_month} 컬럼을 찾지 못했습니다. 재고는 이전 값을 유지하거나 수동 확인이 필요합니다.")
            
        # Fill trade (Rows 35-47 Coated, Rows 52-65 Double Offset)
        # Look up target_ind_month in trade raw sheets
        # For Coated
        r_coated_col = None
        for col in range(ws_r_coated_trade.UsedRange.Columns.Count, 1, -1):
            val = ws_r_coated_trade.Cells(2, col).Value
            if val and str(val).strip() == target_ind_month:
                r_coated_col = col
                break
                
        if r_coated_col:
            ws_m_ind.Cells(35, target_ind_col).Value = target_ind_month
            for offset in range(12):
                ws_m_ind.Cells(36 + offset, target_ind_col).Value = ws_r_coated_trade.Cells(3 + offset, r_coated_col).Value
        else:
            print(f"  ⚠️  Raw '铜版纸进出口量'에서 {target_ind_month} 컬럼을 찾지 못했습니다.")
            
        # For Double Offset
        r_offset_col = None
        for col in range(ws_r_offset_trade.UsedRange.Columns.Count, 1, -1):
            val = ws_r_offset_trade.Cells(2, col).Value
            if val and str(val).strip() == target_ind_month:
                r_offset_col = col
                break
                
        if r_offset_col:
            ws_m_ind.Cells(52, target_ind_col).Value = target_ind_month
            # Double offset has 12 rows (Total Imports to Other Countries Exports)
            for offset in range(12):
                # Row 54 is total imports
                ws_m_ind.Cells(54 + offset, target_ind_col).Value = ws_r_offset_trade.Cells(3 + offset, r_offset_col).Value
        else:
            print(f"  ⚠️  Raw '双胶纸进出口量'에서 {target_ind_month} 컬럼을 찾지 못했습니다.")
            
        # Copy formulas for Summary table (Rows 4-17)
        for r in range(4, 18):
            ws_m_ind.Cells(r, target_ind_col).Formula = ws_m_ind.Cells(r, target_ind_col - 1).Formula
            
        # ─────────────────────────────────────────────────────────────
        # 2C. 수출량 raw 시트 업데이트
        # ─────────────────────────────────────────────────────────────
        ws_m_export = wb_master.Sheets("수출량 raw")
        
        # Find raw sheet named like "2026年4月出口分国别数据"
        raw_export_sheet_name = None
        for name in raw_sheet_names:
            if "出口分国别" in name:
                raw_export_sheet_name = name
                break
                
        if not raw_export_sheet_name:
            print("  [Error] 원본 파일에서 '出口分国别' 시트를 찾지 못했습니다.")
            return
            
        print(f"  - 발견된 국가별 수출 시트: {raw_export_sheet_name}")
        ws_r_export = wb_raw.Sheets(raw_export_sheet_name)
        
        # Parse year/month from the sheet name (e.g. 2026年4月)
        target_y, target_m = parse_date_from_filename(raw_export_sheet_name)
        if not target_y or not target_m:
            # Fallback to indicators month
            if len(date_parts) == 2:
                target_y, target_m = int(date_parts[0]), int(date_parts[1])
            else:
                target_y, target_m = 2026, 4
                
        print(f"  - 국가별 수출 데이터 대상 월: {target_y}년 {target_m}월")
        
        # Load Vlookup translations from master
        ws_vlookup = wb_master.Sheets("Vlookup")
        vlookup_cnt = ws_vlookup.UsedRange.Rows.Count
        translations = {} # cn_name -> (region, kr_name)
        for r in range(1, vlookup_cnt + 1):
            cn = ws_vlookup.Cells(r, 1).Value
            rg = ws_vlookup.Cells(r, 2).Value
            kr = ws_vlookup.Cells(r, 3).Value
            if cn:
                translations[str(cn).strip()] = (str(rg).strip() if rg else "기타", str(kr).strip() if kr else str(cn).strip())
                
        # Read raw export entries from columns A & B (Coated paper)
        r_exp_cnt = ws_r_export.UsedRange.Rows.Count
        new_entries = []
        for r in range(5, r_exp_cnt + 1): # Row 5 is first data row
            cn_country = ws_r_export.Cells(r, 1).Value
            qty = ws_r_export.Cells(r, 2).Value
            if cn_country:
                cn_country_str = str(cn_country).strip()
                if cn_country_str in ["总计", "合计", "Total"]:
                    break
                if qty is not None and str(qty).strip() != "":
                    try:
                        qty_val = float(qty)
                        if qty_val > 0:
                            new_entries.append((cn_country_str, qty_val))
                    except:
                        pass
                        
        print(f"  - 원본 국가별 수출 항목 추출 완료: {len(new_entries)}개")
        
        # Delete existing entries in master sheet '수출량 raw' for target_y, target_m
        # Delete existing entries in master sheet '수출량 raw' for target_y, target_m
        m_exp_cnt = ws_m_export.UsedRange.Rows.Count
        delete_cnt = 0
        if m_exp_cnt > 1:
            y_vals = ws_m_export.Range(ws_m_export.Cells(1, 1), ws_m_export.Cells(m_exp_cnt, 1)).Value
            m_vals = ws_m_export.Range(ws_m_export.Cells(1, 2), ws_m_export.Cells(m_exp_cnt, 2)).Value
            
            delete_rows = []
            for r in range(m_exp_cnt, 1, -1):
                y_val = y_vals[r-1][0]
                m_val = m_vals[r-1][0]
                try:
                    if y_val is not None and m_val is not None:
                        if int(float(y_val)) == int(target_y) and int(float(m_val)) == int(target_m):
                            delete_rows.append(r)
                except:
                    pass
            for r in delete_rows:
                ws_m_export.Rows(r).Delete()
                delete_cnt += 1
        if delete_cnt > 0:
            print(f"  - 기존 {target_y}년 {target_m}월 데이터 {delete_cnt}개 행 삭제 완료 (중복 방지).")
            
        # Append new translated entries
        curr_row = ws_m_export.UsedRange.Rows.Count + 1
        appended_cnt = 0
        missing_translations = []  # 미매핑 국가 추적
        for cn_country, qty in new_entries:
            # Look up translation
            if cn_country in translations:
                region, kr_country = translations[cn_country]
            else:
                region = "기타"
                kr_country = cn_country  # 임시: 중국어 이름 그대로 사용
                missing_translations.append(cn_country)
                
            # Write to cells
            ws_m_export.Cells(curr_row, 1).Value = target_y
            ws_m_export.Cells(curr_row, 2).Value = target_m
            ws_m_export.Cells(curr_row, 3).Value = cn_country
            ws_m_export.Cells(curr_row, 4).Value = region
            ws_m_export.Cells(curr_row, 5).Value = kr_country
            ws_m_export.Cells(curr_row, 6).Value = qty
            
            curr_row += 1
            appended_cnt += 1

        # 미매핑 국가 요약 출력 및 Vlookup 시트 자동 추가
        if missing_translations:
            print(f"\n  ⚠️  [미매핑 국가 발견] Vlookup 시트에 없는 국가 {len(missing_translations)}개:")
            for cn in missing_translations:
                print(f"      - '{cn}' → 한국어 매핑 필요")
            print(f"  → Vlookup 시트에 미매핑 국가를 임시 추가합니다 (지역='기타', 한국어명=중국어명).")
            print(f"  → 추후 Vlookup 시트를 열어 직접 한국어명과 지역을 입력해 주세요.\n")
            if not DRY_RUN:
                vl_row = ws_vlookup.UsedRange.Rows.Count + 1
                existing_cn = set(translations.keys())
                added_to_vlookup = 0
                for cn in missing_translations:
                    if cn not in existing_cn:
                        ws_vlookup.Cells(vl_row, 1).Value = cn
                        ws_vlookup.Cells(vl_row, 2).Value = "기타"
                        ws_vlookup.Cells(vl_row, 3).Value = cn  # 한국어명 임시 = 중국어명
                        vl_row += 1
                        added_to_vlookup += 1
                        existing_cn.add(cn)
                if added_to_vlookup > 0:
                    print(f"  ✅ Vlookup 시트에 {added_to_vlookup}개 미매핑 국가 임시 추가 완료 (직접 수정 필요).")
            else:
                print(f"  [Dry-run] Vlookup 추가를 건너뜁니다.")

        print(f"  - 새 번역 데이터 {appended_cnt}개 행 추가 완료.")
        
        if not DRY_RUN:
            wb_master.Save()
            print("  ✅ 중국 월간 실적 마스터 업데이트 완료!")
        else:
            print("  [Dry-run] 중국 월간 실적 마스터 저장을 건너뜁니다.")
    finally:
        close_wb_safe(wb_raw, was_open_r)
        close_wb_safe(wb_master, was_open_m)

def update_korea_monthly_master(excel, raw_korea_path):
    """Updates the Korean monthly master workbook with filtered rows for the latest month."""
    master_path = MASTER_KR_PATH
    print(f"\n[Part 3] 한국 월간 실적 마스터 업데이트")
    print(f"  - 마스터 파일: {os.path.basename(master_path)}")
    print(f"  - 원본 한국 파일: {os.path.basename(raw_korea_path)}")
    
    # Parse target month from filename
    raw_filename = os.path.basename(raw_korea_path)
    target_y, target_m = parse_date_from_filename(raw_filename)
    
    if not target_y or not target_m:
        print("  [Error] 원본 파일 이름에서 연도와 월을 추출하지 못했습니다.")
        return
        
    # Format target strings for master matching (e.g. 2026 -> 26년, 4 -> 4월)
    year_str = f"{target_y % 100}년"
    month_str = f"{target_m}월"
    ym_key = f"{year_str} {month_str}" # e.g. "26년 4월"
    
    print(f"  - 업데이트 대상 연도월: '{ym_key}'")
    
    # Open workbooks
    wb_master, was_open_m = open_wb_safe(excel, master_path, read_only=False)
    try:
        excel.Calculation = -4135 # xlCalculationManual = -4135
    except Exception as e:
        print(f"  [Warning] Calculation mode change failed: {e}")
    wb_raw, was_open_r = open_wb_safe(excel, raw_korea_path, read_only=True)
    
    try:
        # Find raw sheet '한국산BD ' (with trailing space)
        ws_names = [s.Name for s in wb_raw.Sheets]
        bd_sheet_name = None
        for name in ws_names:
            if "한국산BD" in name:
                bd_sheet_name = name
                break
                
        if not bd_sheet_name:
            print("  [Error] 원본 파일에서 '한국산BD' 시트를 찾지 못했습니다.")
            return
            
        ws_r_bd = wb_raw.Sheets(bd_sheet_name)
        ws_m_bd = wb_master.Sheets("한국산_Raw")
        
        # Read raw rows that match target year & month
        r_bd_cnt = ws_r_bd.UsedRange.Rows.Count
        raw_vals = ws_r_bd.Range(ws_r_bd.Cells(1, 1), ws_r_bd.Cells(r_bd_cnt, 10)).Value # read first 10 columns
        
        filtered_rows = []
        if raw_vals:
            # Skip header row (row 1)
            for row in raw_vals[1:]:
                # Raw layout: 공장, 회사, 연도, 월, 코드, 국가명, 지역, 수량(MT), 금액($), 단가($/MT)
                y_val = str(row[2]).strip() if row[2] is not None else ""
                m_val = str(row[3]).strip() if row[3] is not None else ""
                if y_val == year_str and m_val == month_str:
                    filtered_rows.append(row)
                    
        print(f"  - 원본 필터링 완료: '{ym_key}' 데이터 {len(filtered_rows)}개 행 추출됨")
        
        if not filtered_rows:
            print(f"  ⚠️  원본 파일 시트 '{bd_sheet_name}'에 '{ym_key}'에 대한 행이 존재하지 않습니다. 업데이트를 취소합니다.")
            return
            
        # Delete existing entries in master sheet '한국산_Raw'
        m_r_count = ws_m_bd.UsedRange.Rows.Count
        delete_cnt = 0
        if m_r_count > 1:
            y_vals = ws_m_bd.Range(ws_m_bd.Cells(1, 4), ws_m_bd.Cells(m_r_count, 4)).Value
            m_vals = ws_m_bd.Range(ws_m_bd.Cells(1, 5), ws_m_bd.Cells(m_r_count, 5)).Value
            
            delete_rows = []
            for r in range(m_r_count, 1, -1):
                y_val_val = y_vals[r-1][0]
                m_val_val = m_vals[r-1][0]
                y_val_str = str(y_val_val).strip() if y_val_val is not None else ""
                m_val_str = str(m_val_val).strip() if m_val_val is not None else ""
                if y_val_str == year_str and m_val_str == month_str:
                    delete_rows.append(r)
                    
            for r in delete_rows:
                ws_m_bd.Rows(r).Delete()
                delete_cnt += 1
                
        if delete_cnt > 0:
            print(f"  - 기존 마스터 '{ym_key}' 데이터 {delete_cnt}개 행 삭제 완료 (중복 방지).")
            
        # Append filtered rows
        curr_row = ws_m_bd.UsedRange.Rows.Count + 1
        appended_cnt = 0
        for row in filtered_rows:
            # Master layout:
            # Col A: 연도월 (f"{연도} {월}")
            # Col B to K: Copy Columns 1 to 10 from raw row (공장, 회사, 연도, 월, 코드, 국가명, 지역, 수량, 금액, 단가)
            ws_m_bd.Cells(curr_row, 1).Value = ym_key
            for c_idx in range(10):
                ws_m_bd.Cells(curr_row, 2 + c_idx).Value = row[c_idx]
            curr_row += 1
            appended_cnt += 1
            
        print(f"  - 새 데이터 {appended_cnt}개 행 추가 완료.")
        
        if not DRY_RUN:
            wb_master.Save()
            print("  ✅ 한국 실적 마스터 업데이트 완료!")
        else:
            print("  [Dry-run] 한국 실적 마스터 저장을 건너뜁니다.")
    finally:
        close_wb_safe(wb_raw, was_open_r)
        close_wb_safe(wb_master, was_open_m)

# ─────────────────────────────────────────────
# 5. Compile Dashboard data.js via Excel COM
# ─────────────────────────────────────────────
def safe_float(v, default=0.0):
    try:
        if v is None or pd.isna(v) or v == "/" or str(v).strip() == "":
            return default
        return round(float(v), 1)
    except:
        return default

def safe_int(v, default=0):
    try:
        if v is None or pd.isna(v) or v == "/" or str(v).strip() == "":
            return default
        return int(float(v))
    except:
        return default

def q_label(m):
    return f"Q{(m-1)//3+1}"

def h_label(m):
    return "H1" if m <= 6 else "H2"

def js_array(name, rows):
    """Converts a Python list of dicts to a JavaScript const declaration."""
    import json

    lines = [f"const {name} = ["]
    for r in rows:
        parts = []
        for k, v in r.items():
            if v is None:
                parts.append(f"{k}:null")
            elif isinstance(v, bool):
                parts.append(f"{k}:{'true' if v else 'false'}")
            elif isinstance(v, str):
                parts.append(f"{k}:{json.dumps(v, ensure_ascii=False)}")
            elif isinstance(v, float):
                parts.append(f"{k}:{v:.1f}" if v != int(v) else f"{k}:{int(v)}")
            else:
                parts.append(f"{k}:{v}")
        lines.append("  {" + ", ".join(parts) + "},")
    lines.append("];")
    return "\n".join(lines)
def inject_js_blocks(data_js_content, updates_dict):
    """Replaces or appends const declarations in data.js."""
    for name, rows in updates_dict.items():
        if rows is None:
            continue
        
        # If it is a dict (like DATA_NEWS), handle it differently than array
        if isinstance(rows, dict):
            new_block = f"const {name} = {json_dumps_custom(rows)};"
        else:
            new_block = js_array(name, rows)
            
        pattern = rf"const\s+{re.escape(name)}\s*=\s*(?:\[.*?\]|\{{.*?\}}\s*);\s*"
        if re.search(pattern, data_js_content, re.DOTALL):
            data_js_content = re.sub(pattern, lambda _: new_block + "\n", data_js_content, flags=re.DOTALL)
            print(f"  - {name} 블록 업데이트 완료")
        else:
            data_js_content += "\n\n" + new_block + "\n"
            print(f"  - {name} 블록 신규 추가 완료")
            
    # Update date string
    today_str = datetime.datetime.now().strftime("%Y.%m.%d")
    data_js_content = re.sub(r'최종 업데이트: [\d.]+', f'최종 업데이트: {today_str}', data_js_content)
    return data_js_content

def json_dumps_custom(obj):
    """Outputs Javascript friendly JSON representation for dicts."""
    import json
    return json.dumps(obj, ensure_ascii=False, indent=2)

# ─────────────────────────────────────────────
# 국가 표준화: 국가_마스터.csv(표준국가명, 지역코드, 별칭)
# 지역 코드는 한국 원본 12개 코드(+13_한국, 기타). 중국·한국 원본 모두 같은 표준명·코드로 바꾼다.
# ─────────────────────────────────────────────
COUNTRY_MASTER_PATH = os.path.join(PIPELINE_DIR, "국가_마스터.csv")
KR_REGION_CODES = {"01_미국", "02_호주", "03_서유럽", "04_동유럽", "05_아프리카", "06_남미", "07_일본", "08_중국",
                   "09_동남아", "10_서남아", "11_중동(기타)", "12_중동(이란)", "13_한국"}
# 마스터에 없는 나라가 새로 나오면 중국 원본 코드를 이 규칙으로 바꾼다
CN_TO_KR_REGION = {
    "01_북미": "01_미국", "02_오세아": "02_호주", "03_서유럽_영국": "03_서유럽", "04_서유럽": "03_서유럽",
    "05_동유럽": "04_동유럽", "05_유럽_기타": "04_동유럽", "06_아프리카_남아공": "05_아프리카", "07_아프리카_기타": "05_아프리카",
    "08_남미_멕시코": "06_남미", "09_남미_기타": "06_남미", "10_중국": "08_중국", "11_몽골": "08_중국", "12_대만": "08_중국",
    "13_동남아_말싱": "09_동남아", "14_동남아_인니": "09_동남아", "15_동남아_베트남": "09_동남아", "16_동남아-태국": "09_동남아",
    "17_동남아_필리핀": "09_동남아", "18_동남아_기타": "09_동남아", "19_인도": "10_서남아", "21_방글라": "10_서남아",
    "21_서남아_2": "10_서남아", "22_중동(기타)": "11_중동(기타)", "23_중동(이란)": "12_중동(이란)", "24_일본": "07_일본",
    "25_대한민국": "13_한국",
}
_COUNTRY_REGION, _COUNTRY_ALIAS, UNKNOWN_COUNTRIES = {}, {}, set()


def _load_country_master():
    import csv
    with open(COUNTRY_MASTER_PATH, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            std = row["표준국가명"].strip()
            _COUNTRY_REGION[std] = row["지역코드"].strip()
            for a in (row.get("별칭") or "").split("|"):
                if a.strip():
                    _COUNTRY_ALIAS[a.strip()] = std


def norm_country(name, src_region):
    """(원본 국가명, 원본 지역) → (표준 국가명, 12개 지역 코드). 마스터에 없으면 원본 지역을 변환해 쓰고 기록한다."""
    if not _COUNTRY_REGION:
        _load_country_master()
    n = str(name).strip() if name is not None else ""
    src = str(src_region).strip() if src_region is not None else ""
    if not n or n.startswith("-2146"):  # 엑셀 #N/A
        n = "미상"
    std = _COUNTRY_ALIAS.get(n, n)
    fallback = src if src in KR_REGION_CODES else CN_TO_KR_REGION.get(src, "기타")
    if std not in _COUNTRY_REGION:
        UNKNOWN_COUNTRIES.add(std)
        return std, fallback
    rg = _COUNTRY_REGION[std]
    return std, (fallback if rg == "기타" else rg)


# ─────────────────────────────────────────────
# 국내시장(3페이지) 원본: DRM 엑셀 3개를 읽기 전용으로 연다
# ─────────────────────────────────────────────
DOMESTIC_START_Y = 2023  # 3사 수출 원본이 24년부터. 전년 비교용으로 23년 내수까지 둔다


def _latest_file(folder, pattern):
    files = [f for f in glob.glob(os.path.join(folder, pattern)) if not os.path.basename(f).startswith("~$")]
    return max(files, key=os.path.getmtime) if files else None


def _ym_parts(year_txt, month_txt):
    """'24년'·'2024년' / '9월' → (2024, 9). 못 읽으면 None."""
    y = re.match(r"^\s*(?:20)?(\d\d)년\s*$", str(year_txt or ""))
    m = re.match(r"^\s*(\d{1,2})월\s*$", str(month_txt or ""))
    return (2000 + int(y.group(1)) if y else None, int(m.group(1)) if m else None)


def _sheet_values(wb, name):
    ws = wb.Worksheets(name)
    ur = ws.UsedRange
    r = ur.Row + ur.Rows.Count - 1
    c = ur.Column + ur.Columns.Count - 1
    return [list(row) for row in ws.Range(ws.Cells(1, 1), ws.Cells(r, c)).Value]


def _ms_series(table, label, col=2):
    """M/S 시트: 2행에 연도(1월 열에만), 3행에 월. label 행의 {YYYY-MM: 톤}."""
    row = next(i for i, r in enumerate(table) if col < len(r) and r[col] and str(r[col]).strip() == label)
    cols, cur = {}, None
    for j, (yt, mt) in enumerate(zip(table[1], table[2])):
        y, _ = _ym_parts(yt, None)
        cur = y or cur
        _, m = _ym_parts(None, mt)
        if cur and m and (cur, m) not in cols:
            cols[(cur, m)] = j
    out = {}
    for (y, m), j in sorted(cols.items()):
        v = table[row][j]
        if y >= DOMESTIC_START_Y and isinstance(v, (int, float)) and v != 0 and abs(v) < 1e9:
            out[f"{y}-{m:02d}"] = round(v, 1)
    return out


def extract_domestic_market(excel):
    """국내 내수(M/S)·3사 도공지 수출·미국향 백상지를 읽어 DATA_DOM_* 블록을 만든다.
    도공지(내수)=아트지만(MFC 제외), 비도공지=백상지만(복사지 제외). 용어는 CONTEXT.md 참고."""
    print("  - 국내시장 데이터 처리 중...")
    f_ms = _latest_file(DOMESTIC_DIR, "인쇄용지*MS*.xlsx")
    f_co = _latest_file(DOMESTIC_DIR, "(인쇄) 한국산 도공지 수출실적_*.xlsx")
    f_us = _latest_file(DOMESTIC_DIR, "(인쇄) 미국 비도공지 수출실적_*.xlsx")
    if not (f_ms and f_co and f_us):
        raise FileNotFoundError(f"국내시장 원본 누락({DOMESTIC_DIR}): MS={f_ms}, 도공지={f_co}, 미국={f_us}")

    sheets = {}
    for path, names in [(f_ms, ["아트지", "백상지"]), (f_co, ["한국산BD ", "1-1 종합"]), (f_us, ["한국산BD"])]:
        wb, was_open = open_wb_safe(excel, path, read_only=True)
        try:
            for n in names:
                sheets[(path, n)] = _sheet_values(wb, n)
        finally:
            close_wb_safe(wb, was_open)

    art, wf = sheets[(f_ms, "아트지")], sheets[(f_ms, "백상지")]
    ms = {
        "art_maker": _ms_series(art, "계[A]"),
        "art_demand": _ms_series(art, "국내 총수요", col=1),
        "wf_maker": _ms_series(wf, "계[A]"),
        "wf_demand": _ms_series(wf, "국내 총수요", col=1),
        "files": {"ms": os.path.basename(f_ms), "coated": os.path.basename(f_co), "us": os.path.basename(f_us)},
    }

    # 3사 도공지 수출: 공장,회사,연도,월,코드,국가명,지역,수량(MT) …
    bd = sheets[(f_co, "한국산BD ")]
    hdr = [str(h).strip() if h else "" for h in bd[0]]
    ix = {k: hdr.index(k) for k in ["회사", "연도", "월", "국가명", "지역", "수량(MT)"]}
    agg = {}
    for r in bd[1:]:
        y, m = _ym_parts(r[ix["연도"]], r[ix["월"]])
        q = r[ix["수량(MT)"]]
        if not (y and m) or not isinstance(q, (int, float)) or q == 0:
            continue
        k = (f"{y}-{m:02d}", str(r[ix["회사"]]).strip(), *norm_country(r[ix["국가명"]], r[ix["지역"]]))
        agg[k] = agg.get(k, 0) + q
    rows_export = [{"ym": k[0], "comp": k[1], "ct": k[2], "rg": k[3], "v": round(v, 1)} for k, v in sorted(agg.items())]

    # 대사: 월별 합계 vs '1-1 종합' 총합계(7번째 열)
    month_tot = {}
    for r in rows_export:
        month_tot[r["ym"]] = month_tot.get(r["ym"], 0) + r["v"]
    cur_y, checked, diffs = None, 0, []
    for r in sheets[(f_co, "1-1 종합")][7:]:
        y, _ = _ym_parts(r[1], None)
        cur_y = y or cur_y
        _, m = _ym_parts(None, r[2])
        if not (cur_y and m) or not isinstance(r[6], (int, float)):
            continue
        checked += 1
        k = f"{cur_y}-{m:02d}"
        if abs(month_tot.get(k, 0) - r[6]) > 1:
            diffs.append((k, round(month_tot.get(k, 0), 1), round(r[6], 1)))
    print(f"    · 3사 도공지 월합계 대사: {checked}개월 확인, 불일치 {len(diffs)}건 {diffs[:3]}")

    # 미국향 비도공: 지류 '백상지'만(Box 복사지 제외)
    ub = sheets[(f_us, "한국산BD")]
    uh = [str(h).strip() if h else "" for h in ub[0]]
    ui = {k: uh.index(k) for k in ["회사", "통관량", "연도", "월", "지류"]}
    us = {}
    for r in ub[1:]:
        y, m = _ym_parts(r[ui["연도"]], r[ui["월"]])
        q = r[ui["통관량"]]
        if not (y and m) or str(r[ui["지류"]]).strip() != "백상지" or not isinstance(q, (int, float)):
            continue
        k = (f"{y}-{m:02d}", str(r[ui["회사"]]).strip())
        us[k] = us.get(k, 0) + q
    rows_us = [{"ym": k[0], "comp": k[1], "v": round(v, 1)} for k, v in sorted(us.items())]

    print(f"    · 내수 최신 {max(ms['art_maker'])}, 3사 수출 최신 {max(month_tot)} ({len(rows_export)}행), 미국향 최신 {rows_us[-1]['ym'] if rows_us else '-'}")
    return {"DATA_DOM_MS": ms, "DATA_DOM_EXPORT": rows_export, "DATA_DOM_US": rows_us}

def extract_and_sync_dashboard_data(excel, pdf_summary_item=None):
    """Extracts master data using Excel COM and updates data.js."""
    print(f"\n[Part 5] 대시보드 data.js 파일 갱신")
    
    kr_master_path = MASTER_KR_PATH
    cn_master_path = MASTER_CN_PATH
    
    if not os.path.exists(kr_master_path) or not os.path.exists(cn_master_path):
        print("  [Error] 마스터 파일이 유실되었습니다. 갱신을 취소합니다.")
        return
        
    # Open workbooks
    wb_kr, was_open_kr = open_wb_safe(excel, kr_master_path, read_only=True)
    try:
        excel.Calculation = -4135 # xlCalculationManual = -4135
    except Exception as e:
        print(f"  [Warning] Calculation mode change failed: {e}")
    wb_cn, was_open_cn = open_wb_safe(excel, cn_master_path, read_only=True)
    
    all_js_updates = {}
    
    try:
        # ─────────────────────────────────────────────────────────────
        # 5A. 한국산 데이터 파싱 및 가공 (DATA_KR_...)
        # ─────────────────────────────────────────────────────────────
        print("  - 한국산 실적 데이터 처리 중...")
        ws_kr_raw = wb_kr.Sheets("한국산_Raw")
        used_kr = ws_kr_raw.UsedRange
        kr_r_cnt = used_kr.Rows.Count
        
        # Read entire sheet in memory
        kr_vals = ws_kr_raw.Range(ws_kr_raw.Cells(1, 1), ws_kr_raw.Cells(kr_r_cnt, 11)).Value
        
        # Load into pandas for clean grouping
        df_kr = pd.DataFrame(kr_vals[1:], columns=kr_vals[0])
        
        # Clean data
        df_kr = df_kr.dropna(subset=['연도', '월', '회사', '수량(MT)'])
        
        # Map columns using fallback names
        col_year = '연도'
        col_month = '월'
        col_comp = '회사'
        col_region = '지역'
        col_country = '국가명'
        col_volume = '수량(MT)'
        col_amount = '금액($)'
        
        # Clean types
        def clean_year(y):
            # "26년" or 2026.0 -> 2026
            if isinstance(y, str):
                y_num = re.sub(r'\D', '', y)
                if y_num:
                    return 2000 + int(y_num) if len(y_num) == 2 else int(y_num)
            try:
                val = int(float(y))
                return 2000 + val if val < 100 else val
            except:
                return 2018
                
        def clean_month(m):
            # "4월" or 4.0 -> 4
            if isinstance(m, str):
                m_num = re.sub(r'\D', '', m)
                if m_num: return int(m_num)
            try:
                return int(float(m))
            except:
                return 1
                
        df_kr['y'] = df_kr[col_year].apply(clean_year)
        df_kr['m'] = df_kr[col_month].apply(clean_month)
        df_kr['v'] = df_kr[col_volume].apply(safe_float)
        df_kr['amt'] = df_kr[col_amount].apply(safe_float)
        df_kr['comp'] = df_kr[col_comp].apply(lambda x: str(x).strip() if x else "")
        # 국가명·지역은 국가_마스터.csv 기준으로 통일(중국산과 같은 표준명·같은 12개 지역 코드)
        normed = [norm_country(c, r) for c, r in zip(df_kr[col_country], df_kr[col_region])]
        df_kr['c'] = [n[0] for n in normed]
        df_kr['rg'] = [n[1] for n in normed]
        
        # Group 1: 회사별 월합계 (DATA_KR_EXPORT)
        g_exp = df_kr.groupby(['y', 'm', 'comp']).agg({'v': 'sum', 'amt': 'sum'}).reset_index()
        rows_kr_export = []
        for _, r in g_exp.iterrows():
            y, m, comp, v, amt = int(r['y']), int(r['m']), r['comp'], r['v'], r['amt']
            fob = round(amt / v, 1) if v > 0 else 0.0
            rows_kr_export.append({
                "y": y, "m": m, "q": q_label(m), "h": h_label(m),
                "comp": comp, "v": safe_int(v), "amt": safe_int(amt), "fob": fob
            })
            
        # Group 2: 지역별 (DATA_KR_REGION)
        g_reg = df_kr.groupby(['y', 'm', 'comp', 'rg']).agg({'v': 'sum', 'amt': 'sum'}).reset_index()
        rows_kr_region = []
        for _, r in g_reg.iterrows():
            y, m, comp, rg, v, amt = int(r['y']), int(r['m']), r['comp'], r['rg'], r['v'], r['amt']
            fob = round(amt / v, 1) if v > 0 else 0.0
            rows_kr_region.append({
                "y": y, "m": m, "q": q_label(m), "h": h_label(m),
                "comp": comp, "rg": rg, "v": safe_int(v), "amt": safe_int(amt), "fob": fob
            })
            
        # Group 3: 국가별 (DATA_KR_COUNTRY)
        g_coun = df_kr.groupby(['y', 'm', 'comp', 'c', 'rg']).agg({'v': 'sum', 'amt': 'sum'}).reset_index()
        rows_kr_country = []
        for _, r in g_coun.iterrows():
            y, m, comp, c, rg, v, amt = int(r['y']), int(r['m']), r['comp'], r['c'], r['rg'], r['v'], r['amt']
            fob = round(amt / v, 1) if v > 0 else 0.0
            rows_kr_country.append({
                "y": y, "m": m, "q": q_label(m), "h": h_label(m),
                "comp": comp, "c": c, "rg": rg, "v": safe_int(v), "amt": safe_int(amt), "fob": fob
            })
            
        # Group 4: 연간 (DATA_KR_YEARLY)
        g_year = df_kr.groupby(['y', 'comp']).agg({'v': 'sum', 'amt': 'sum'}).reset_index()
        rows_kr_yearly = []
        for _, r in g_year.iterrows():
            y, comp, v, amt = int(r['y']), r['comp'], r['v'], r['amt']
            fob = round(amt / v, 1) if v > 0 else 0.0
            rows_kr_yearly.append({
                "y": y, "comp": comp, "v": safe_int(v), "amt": safe_int(amt), "fob": fob
            })
            
        # Group 5: 연간 합계 (DATA_KR_TOTAL)
        g_total = df_kr.groupby(['y']).agg({'v': 'sum', 'amt': 'sum'}).reset_index()
        rows_kr_total = []
        for _, r in g_total.iterrows():
            y, v, amt = int(r['y']), r['v'], r['amt']
            fob = round(amt / v, 1) if v > 0 else 0.0
            rows_kr_total.append({
                "y": y, "v": safe_int(v), "amt": safe_int(amt), "fob": fob
            })
            
        # Sort rows
        rows_kr_export.sort(key=lambda x: (x["y"], x["m"], x["comp"]))
        rows_kr_region.sort(key=lambda x: (x["y"], x["m"], x["comp"], x["rg"]))
        rows_kr_country.sort(key=lambda x: (x["y"], x["m"], x["comp"], x["c"]))
        rows_kr_yearly.sort(key=lambda x: (x["y"], x["comp"]))
        rows_kr_total.sort(key=lambda x: x["y"])
        
        all_js_updates.update({
            "DATA_KR_EXPORT": rows_kr_export,
            "DATA_KR_REGION": rows_kr_region,
            "DATA_KR_COUNTRY": rows_kr_country,
            "DATA_KR_YEARLY": rows_kr_yearly,
            "DATA_KR_TOTAL": rows_kr_total
        })
        
        # ─────────────────────────────────────────────────────────────
        # 5B. 중국산 도착국가별 데이터 파싱 (DATA_EXPORT)
        # ─────────────────────────────────────────────────────────────
        print("  - 중국산 실적 데이터 처리 중...")
        ws_cn_export = wb_cn.Sheets("수출량 raw")
        used_cn_exp = ws_cn_export.UsedRange
        cn_exp_cnt = used_cn_exp.Rows.Count
        
        cn_exp_vals = ws_cn_export.Range(ws_cn_export.Cells(1, 1), ws_cn_export.Cells(cn_exp_cnt, 6)).Value
        
        rows_cn_export = []
        if cn_exp_vals:
            # Row 1 is header. 열 순서: 년도, 월, 지역, 도착국가(한국어), 도착국가(중국어), 중량(MT)
            # (261105 수정: 이전 코드가 지역=row[3], 국가=row[4]로 한 칸 밀려 읽어 중국산이 전부 '기타' 권역으로 잡혔음)
            hdr = [str(h).strip() if h else "" for h in cn_exp_vals[0]]
            if len(hdr) < 4 or hdr[2] != "지역" or "한국어" not in hdr[3]:
                print(f"  [Warning] '수출량 raw' 헤더가 예상과 다릅니다: {hdr}")
            # 일부 달(예: 26.4월)은 옛 순서(년도, 월, 중국어, 지역, 한국어)로 들어 있어 행마다 지역 코드 위치로 판단
            is_rg = lambda v: bool(re.match(r"^\d\d_", str(v or "").strip()))
            for row in cn_exp_vals[1:]:
                y_val = row[0]
                m_val = row[1]
                if not is_rg(row[2]) and is_rg(row[3]):
                    region, kr_country = row[3], row[4]
                else:
                    region, kr_country = row[2], row[3]
                qty = row[5]
                if y_val and m_val and kr_country:
                    c_std, rg_std = norm_country(kr_country, region)
                    rows_cn_export.append({
                        "y": safe_int(y_val),
                        "m": safe_int(m_val),
                        "c": c_std,
                        "_rg": rg_std,
                        "v": safe_int(qty)
                    })
        rows_cn_export.sort(key=lambda x: (x["y"], x["m"], x["c"]))
        all_js_updates["DATA_EXPORT"] = rows_cn_export
        
        # ─────────────────────────────────────────────────────────────
        # 5C. 중국 시장 실적 지표 파싱 (DATA_INDICATORS)
        # ─────────────────────────────────────────────────────────────
        print("  - 실적지표 데이터 처리 중...")
        ws_m_ind = wb_cn.Sheets("주요지표")
        used_ind = ws_m_ind.UsedRange
        ind_c_cnt = used_ind.Columns.Count
        
        rows_indicators = []
        for col in range(3, ind_c_cnt + 1):
            p_val = ws_m_ind.Cells(3, col).Value
            if not p_val or not str(p_val).strip():
                continue
            p_str = str(p_val).strip()
            
            # Parse year/month
            y, m = parse_excel_date(p_str)
            if not y:
                continue
                
            cb = "1-2" in p_str or "1~2" in p_str
            
            pro_val = ws_m_ind.Cells(4, col).Value
            inv_val = ws_m_ind.Cells(5, col).Value
            exp_val = ws_m_ind.Cells(6, col).Value
            imp_val = ws_m_ind.Cells(7, col).Value
            dom_val = ws_m_ind.Cells(8, col).Value

            # 기간 라벨만 있고 값이 비어 있는 미래 월(엑셀 빈 칸)은 건너뜀
            if all(safe_float(v, None) is None for v in (pro_val, inv_val, exp_val, imp_val, dom_val)):
                continue

            rows_indicators.append({
                "p": p_str,
                "y": y,
                "m": m,
                "cb": cb,
                "pro": safe_float(pro_val, None),
                "inv": safe_float(inv_val, None),
                "exp": safe_float(exp_val, None),
                "imp": safe_float(imp_val, None),
                "dom": safe_float(dom_val, None)
            })
            
        rows_indicators.sort(key=lambda x: (x["y"], x["m"]))
        all_js_updates["DATA_INDICATORS"] = rows_indicators
        
        # ─────────────────────────────────────────────────────────────
        # 5D. UMPaper 가격 인덱스 파싱 (DATA_PRICE)
        # ─────────────────────────────────────────────────────────────
        print("  - 가격지수 데이터 처리 중...")
        ws_m_idx = wb_cn.Sheets('UMPaper文化纸价格指数 (가격 Index)')
        used_idx = ws_m_idx.UsedRange
        idx_c_cnt = used_idx.Columns.Count
        
        prices_list = []
        usd_list = []
        
        # Row 3 to 11 represent prices
        # Row 14 is domestic USD, Row 15 is expected import USD
        # Columns start at 4 (2024/7)
        for col in range(4, idx_c_cnt + 1):
            h_val = ws_m_idx.Cells(2, col).Value
            if not h_val or not str(h_val).strip():
                continue
            h_str = str(h_val).strip()
            
            # Parse year/month
            y, m = parse_excel_date(h_str)
            if not y:
                continue
                
            # Prices (Row 3-11)
            # Row 3: 코트지 광동
            p_c_gd = ws_m_idx.Cells(3, col).Value
            if p_c_gd and str(p_c_gd).strip() != "":
                prices_list.append({"t": "코트지", "g": "157gsm", "r": "광동", "y": y, "m": m, "p": safe_float(p_c_gd)})
            # Row 4: 코트지 장절호
            p_c_zj = ws_m_idx.Cells(4, col).Value
            if p_c_zj and str(p_c_zj).strip() != "":
                prices_list.append({"t": "코트지", "g": "157gsm", "r": "장절호", "y": y, "m": m, "p": safe_float(p_c_zj)})
            # Row 5: 코트지 북경
            p_c_bj = ws_m_idx.Cells(5, col).Value
            if p_c_bj and str(p_c_bj).strip() != "":
                prices_list.append({"t": "코트지", "g": "157gsm", "r": "북경", "y": y, "m": m, "p": safe_float(p_c_bj)})
            # Row 6: 코트지 하남
            p_c_hn = ws_m_idx.Cells(6, col).Value
            if p_c_hn and str(p_c_hn).strip() != "":
                prices_list.append({"t": "코트지", "g": "157gsm", "r": "하남", "y": y, "m": m, "p": safe_float(p_c_hn)})
            # Row 7: 코트지 사천
            p_c_sc = ws_m_idx.Cells(7, col).Value
            if p_c_sc and str(p_c_sc).strip() != "":
                prices_list.append({"t": "코트지", "g": "157gsm", "r": "사천", "y": y, "m": m, "p": safe_float(p_c_sc)})
            # Row 8: 옵셋지 광동
            p_o_gd = ws_m_idx.Cells(8, col).Value
            if p_o_gd and str(p_o_gd).strip() != "":
                prices_list.append({"t": "옵셋지", "g": "70/80gsm", "r": "광동", "y": y, "m": m, "p": safe_float(p_o_gd)})
            # Row 9: 옵셋지 장절호
            p_o_zj = ws_m_idx.Cells(9, col).Value
            if p_o_zj and str(p_o_zj).strip() != "":
                prices_list.append({"t": "옵셋지", "g": "70/80gsm", "r": "장절호", "y": y, "m": m, "p": safe_float(p_o_zj)})
            # Row 10: 옵셋지 북경
            p_o_bj = ws_m_idx.Cells(10, col).Value
            if p_o_bj and str(p_o_bj).strip() != "":
                prices_list.append({"t": "옵셋지", "g": "70/80gsm", "r": "북경", "y": y, "m": m, "p": safe_float(p_o_bj)})
            # Row 11: 옵셋지 하남
            p_o_hn = ws_m_idx.Cells(11, col).Value
            if p_o_hn and str(p_o_hn).strip() != "":
                prices_list.append({"t": "옵셋지", "g": "70/80gsm", "r": "하남", "y": y, "m": m, "p": safe_float(p_o_hn)})
                
            # USD Exchange / import (Row 14, 15)
            d_val = ws_m_idx.Cells(14, col).Value
            e_val = ws_m_idx.Cells(15, col).Value
            if d_val or e_val:
                usd_list.append({
                    "y": y,
                    "m": m,
                    "d": safe_float(d_val),
                    "e": safe_float(e_val)
                })
                
        all_js_updates["DATA_PRICE"] = {
            "prices": prices_list,
            "usd": usd_list
        }
        
        # ─────────────────────────────────────────────────────────────
        # 5E. 뉴스 피드 추가 (DATA_NEWS)
        # ─────────────────────────────────────────────────────────────
        if pdf_summary_item:
            print("  - 요약 뉴스 업데이트 처리 중...")
            # Read existing DATA_NEWS from data.js
            with open(DATA_JS_PATH, "r", encoding="utf-8") as f:
                content = f.read()
                
            match = re.search(r"const\s+DATA_NEWS\s*=\s*(.*?);\s*\n", content, re.DOTALL)
            if match:
                import json
                try:
                    # Clean up javascript object to parse as JSON
                    js_obj_str = match.group(1).strip()
                    # A very simple parsing fallback
                    # If it contains JS variables or unquoted keys, json.loads will fail.
                    # We can construct the news feed by keeping the structure.
                    # The easiest way: build the new DATA_NEWS object
                    # Format: { "last_updated": "YYYY-MM-DD", "items": [ new_item, ...old_items ] }
                    
                    # We can do regex replacement or clean JSON parsing.
                    # Let's write a robust parser/writer:
                    # Parse old items using regex
                    old_items = []
                    item_matches = re.findall(r"\{\s*date:\s*\"(.*?)\",\s*title:\s*\"(.*?)\",\s*source:\s*\"(.*?)\",\s*url:\s*\"(.*?)\",\s*keywords:\s*\[(.*?)\],\s*summary:\s*\"(.*?)\"\s*\}", js_obj_str, re.DOTALL)
                    for date, title, source, url, kw_str, summary in item_matches:
                        kws = [k.strip().strip('"').strip("'") for k in kw_str.split(",")]
                        old_items.append({
                            "date": date,
                            "title": title,
                            "source": source,
                            "url": url,
                            "keywords": kws,
                            "summary": summary
                        })
                        
                    # Prepend new item
                    new_items = [pdf_summary_item]
                    # Filter out duplicates (based on title)
                    seen_titles = {pdf_summary_item["title"]}
                    for item in old_items:
                        if item["title"] not in seen_titles:
                            new_items.append(item)
                            seen_titles.add(item["title"])
                            
                    all_js_updates["DATA_NEWS"] = {
                        "last_updated": datetime.datetime.now().strftime("%Y-%m-%d"),
                        "items": new_items[:15] # Keep last 15 items
                    }
                except Exception as ex:
                    print(f"  ⚠️  기존 뉴스 피드 파싱 실패 ({ex}), 새 뉴스 피드로 대체합니다.")
                    all_js_updates["DATA_NEWS"] = {
                        "last_updated": datetime.datetime.now().strftime("%Y-%m-%d"),
                        "items": [pdf_summary_item]
                    }
            else:
                all_js_updates["DATA_NEWS"] = {
                    "last_updated": datetime.datetime.now().strftime("%Y-%m-%d"),
                    "items": [pdf_summary_item]
                }
                
        # ─────────────────────────────────────────────────────────────
        # 5E-2. 국내시장(3페이지) 데이터. 실패해도 기존 data.js 값을 유지하고 나머지는 계속 진행
        # ─────────────────────────────────────────────────────────────
        try:
            all_js_updates.update(extract_domestic_market(excel))
        except Exception as ex:
            print(f"  ⚠️  국내시장 데이터 갱신 실패 ({ex}). 기존 DATA_DOM_* 값을 유지합니다.")
        if UNKNOWN_COUNTRIES:
            print(f"  ⚠️  국가_마스터.csv에 없는 국가 {len(UNKNOWN_COUNTRIES)}개(원본 지역으로 임시 분류): {sorted(UNKNOWN_COUNTRIES)}")

        # ─────────────────────────────────────────────────────────────
        # 5F. DATA_META 생성 (기준월 정보)
        # ─────────────────────────────────────────────────────────────
        # Determine latest periods from extracted data
        def _latest_period(rows, y_key='y', m_key='m'):
            """Returns (year, month) of the latest record in rows."""
            if not rows:
                return None, None
            latest = max(rows, key=lambda r: (r.get(y_key, 0), r.get(m_key, 0)))
            return latest.get(y_key), latest.get(m_key)

        cn_price_y, cn_price_m = _latest_period(prices_list)
        cn_monthly_y, cn_monthly_m = _latest_period(rows_indicators)
        kr_y, kr_m = _latest_period(rows_kr_export)

        data_meta = {
            "last_run": datetime.datetime.now().strftime("%Y-%m-%d"),
            "cn_price_latest": f"{cn_price_y}-{cn_price_m:02d}" if cn_price_y else "",
            "cn_monthly_latest": f"{cn_monthly_y}-{cn_monthly_m:02d}" if cn_monthly_y else "",
            "kr_monthly_latest": f"{kr_y}-{kr_m:02d}" if kr_y else "",
            "version": "1.1"
        }
        all_js_updates["DATA_META"] = data_meta
        print(f"  - DATA_META 생성: 중국가격={data_meta['cn_price_latest']}, 중국실적={data_meta['cn_monthly_latest']}, 한국실적={data_meta['kr_monthly_latest']}")

        # ─────────────────────────────────────────────────────────────
        # 5G. data.js 파일에 갱신 데이터 주입
        # ─────────────────────────────────────────────────────────────
        print("  - data.js 파일 쓰기 작업 중...")
        with open(DATA_JS_PATH, "r", encoding="utf-8") as f:
            js_content = f.read()
            
        # Create a backup of data.js
        if not DRY_RUN:
            backup_path = DATA_JS_PATH + f".backup_{datetime.datetime.now().strftime('%y%m%d_%H%M')}"
            with open(backup_path, "w", encoding="utf-8") as f_bak:
                f_bak.write(js_content)
            print(f"  - data.js 백업 저장 완료: {os.path.basename(backup_path)}")
            
            js_content = inject_js_blocks(js_content, all_js_updates)
            
            with open(DATA_JS_PATH, "w", encoding="utf-8") as f:
                f.write(js_content)
            print("  ✅ data.js 업데이트 완료!")
        else:
            print("  [Dry-run] data.js 백업 및 쓰기를 건너뜁니다. (새 데이터 요약):")
            for name, rows in all_js_updates.items():
                if isinstance(rows, list):
                    print(f"    - {name}: {len(rows)} 행")
                elif isinstance(rows, dict):
                    print(f"    - {name}: {list(rows.keys())} 키의 딕셔너리")
        
    finally:
        wb_kr.Close(False)
        wb_cn.Close(False)

# ─────────────────────────────────────────────
# 6. Auto Backup
# ─────────────────────────────────────────────
def auto_backup(selected_files, phase="pre"):
    """Backs up master excels, raw source files, and data.js to a timestamped folder.
    
    Args:
        selected_files: dict with keys 'cn_weekly', 'cn_monthly', 'cn_pdf', 'kr_monthly'
        phase: 'pre' (before run) or 'post' (after run)
    """
    import shutil
    now = datetime.datetime.now()
    month_label = now.strftime("%Y-%m")
    backup_root = os.path.join(WORKSPACE_DIR, "Dashboard_Backup", month_label)
    backup_dir = os.path.join(backup_root, f"{phase}_run_{now.strftime('%H%M')}")

    os.makedirs(backup_dir, exist_ok=True)
    print(f"\n[Backup/{phase}] 백업 폴더: {backup_dir}")

    # 1. Master Excel backup
    master_dir = os.path.join(backup_dir, "master_excels")
    os.makedirs(master_dir, exist_ok=True)
    for src in MASTER_PATHS:
        fn = os.path.basename(src)
        if os.path.exists(src):
            try:
                shutil.copy2(src, os.path.join(master_dir, fn))
                print(f"  - 마스터 백업: {fn}")
            except Exception as e:
                print(f"  ⚠️  마스터 백업 실패 ({fn}): {e}")

    # 2. Raw source files backup (선택된 파일만)
    if phase == "pre" and selected_files:
        raw_dir = os.path.join(backup_dir, "raw_files")
        os.makedirs(raw_dir, exist_ok=True)
        for key, path in selected_files.items():
            if path and os.path.exists(path):
                try:
                    shutil.copy2(path, os.path.join(raw_dir, os.path.basename(path)))
                    print(f"  - 원본 파일 백업: {os.path.basename(path)}")
                except Exception as e:
                    print(f"  ⚠️  원본 파일 백업 실패: {e}")

    # 3. data.js snapshot
    if os.path.exists(DATA_JS_PATH):
        data_dir = os.path.join(backup_dir, "data_js")
        os.makedirs(data_dir, exist_ok=True)
        snap_name = f"data_{now.strftime('%Y%m%d_%H%M')}_{phase}.js"
        try:
            shutil.copy2(DATA_JS_PATH, os.path.join(data_dir, snap_name))
            print(f"  - data.js 스냅샷 저장: {snap_name}")
        except Exception as e:
            print(f"  ⚠️  data.js 스냅샷 실패: {e}")

    print(f"  ✅ [{phase}] 백업 완료")
    return backup_dir


def commit_and_push_to_github():
    """Stages updated files and Excel master back-ups, commits, and pushes to GitHub Pages."""
    print("\n[Part 6] 깃허브(서버) 업로드 자동화")
    
    # Backup the master Excel files by copying them into 'master_excels/' in the dashboard repo
    backup_dir = os.path.join(DASHBOARD_DIR, "master_excels")
    os.makedirs(backup_dir, exist_ok=True)
    
    master_files = [os.path.basename(p) for p in MASTER_PATHS]

    import shutil
    for src in MASTER_PATHS:
        fn = os.path.basename(src)
        dst = os.path.join(backup_dir, fn)
        if os.path.exists(src):
            try:
                if not DRY_RUN:
                    shutil.copy2(src, dst)
                    print(f"  - 마스터 백업 파일 복사 완료: {fn} -> master_excels/")
                else:
                    print(f"  [Dry-run] 마스터 백업 파일 복사를 건너뜁니다: {fn}")
            except Exception as e:
                print(f"  ⚠️  마스터 파일 백업 복사 실패 ({fn}): {e}")
                
    def run_git(args):
        res = subprocess.run(["git"] + args, cwd=DASHBOARD_DIR, capture_output=True, text=True)
        return res
        
    if DRY_RUN:
        print("  [Dry-run] Dry-run 모드로 깃허브 푸시를 건너뜁니다.")
        return
        
    # Check git repository
    git_status = run_git(["status"])
    if git_status.returncode != 0:
        print("  - Git 저장소가 감지되지 않아 새로 초기화합니다.")
        run_git(["init"])
        run_git(["branch", "-M", "main"])
        
    # Check remote origin
    git_remote = run_git(["remote", "-v"])
    target_url = "https://github.com/sungsulee1190-star/pnwdashboard"
    if target_url not in git_remote.stdout:
        print(f"  - 원격 저장소 등록: {target_url}")
        if "origin" in git_remote.stdout:
            run_git(["remote", "remove", "origin"])
        run_git(["remote", "add", "origin", target_url])
        
    # Stage files
    run_git(["add", "data.js"])
    for fn in master_files:
        run_git(["add", f"master_excels/{fn}"])
        
    # Check if there are changes to commit
    status_output = run_git(["status", "--porcelain"])
    if not status_output.stdout.strip():
        print("  - 변경된 파일이 없어 깃허브 업로드를 건너뜁니다.")
        return
        
    # git 작성자 정보는 코드에 두지 않는다(공개 저장소). 이 PC의 git 설정을 쓴다
    if run_git(["config", "user.email"]).returncode != 0:
        print("  - [오류] git user.email이 설정되지 않아 업로드를 건너뜁니다. 저장소 폴더에서 git config user.email 을 한 번 설정하세요.")
        return

    # Commit
    today = datetime.datetime.now().strftime("%Y-%m-%d")
    commit_msg = f"data: {today} 자동 업데이트"
    run_git(["commit", "-m", commit_msg])
    print(f"  - 로컬 커밋 완료: '{commit_msg}'")
    
    # Push (실패 시 force push 재시도)
    print("  - 깃허브 푸시(push) 진행 중 (자격 증명이 필요할 수 있습니다)...")
    push_res = run_git(["push", "-u", "origin", "main"])
    if push_res.returncode == 0:
        print("  ✅ 깃허브(서버) 업로드 완료!")
    else:
        # 원격 업데이트가 있을 때 force push
        print("  ⚠️  일반 push 실패, --force로 재시도...")
        push_res2 = run_git(["push", "-u", "origin", "main", "--force"])
        if push_res2.returncode == 0 or "forced update" in push_res2.stderr:
            print("  ✅ 깃허브(서버) 업로드 완료! (force push)")
        else:
            print(f"  ❌ 깃허브 푸시 실패:\n{push_res2.stderr}")

# ─────────────────────────────────────────────
# Main Pipeline
# ─────────────────────────────────────────────
def main():
    print("=" * 70)
    print("인쇄용지 대시보드 경쟁 분석 데이터 자동 갱신 파이프라인")
    print(f"실행 시간: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print("=" * 70)
    
    # 1. Locate Raw Files
    cn_weekly = robust_find_file(RAW_CN_DIR, ["价格数据", "2026"], ".xlsx")
    cn_monthly = robust_find_file(RAW_CN_DIR, ["月报数据", "2026"], ".xlsx")
    cn_pdf = robust_find_file(RAW_CN_DIR, ["市场月报", "2026"], ".pdf")
    kr_monthly = robust_find_file(RAW_KR_DIR, ["수출실적", "2026"], ".xlsx")

    selected_files = {
        "cn_weekly": cn_weekly,
        "cn_monthly": cn_monthly,
        "cn_pdf": cn_pdf,
        "kr_monthly": kr_monthly,
    }

    # Check if files found — 상세 정보 출력 (파일명 + 수정일 + 기준월)
    def _file_info(path):
        if not path:
            return "❌ 미발견"
        mtime = datetime.datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d %H:%M")
        y, m = parse_date_from_filename(os.path.basename(path))
        period = f"{y}년 {m}월" if y and m else "(기준월 미파싱)"
        return f"{os.path.basename(path)}\n      수정일: {mtime} | 기준: {period}"

    print("발견된 최신 원본 파일들:")
    print(f"  - 중국 주간 가격: {_file_info(cn_weekly)}")
    print(f"  - 중국 월간 데이터: {_file_info(cn_monthly)}")
    print(f"  - 중국 시장 리포트 (PDF): {_file_info(cn_pdf)}")
    print(f"  - 한국 월간 데이터: {_file_info(kr_monthly)}")
    
    # 2. PDF 번역·요약은 파이프라인에서 하지 않는다(API 호출 배제). 필요하면 구독 모델로 따로 한다.
    pdf_summary_item = None

    # 2-a. Pre-run backup (마스터 엑셀 + 원본 파일 + data.js 스냅샷)
    if not DRY_RUN:
        auto_backup(selected_files, phase="pre")

    # 3. Excel COM Updates
    excel = None
    try:
        print("\nExcel COM 연결 중...")
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        print("Excel COM 연결 완료.")
        
        # 3A. Update China Weekly
        if cn_weekly:
            update_weekly_price_master(excel, cn_weekly)
            
        # 3B. Update China Monthly
        if cn_monthly:
            update_china_monthly_master(excel, cn_monthly)
            
        # 3C. Update Korea Monthly
        if kr_monthly:
            update_korea_monthly_master(excel, kr_monthly)
            
        # 4. Extract and Sync to Dashboard data.js
        extract_and_sync_dashboard_data(excel, pdf_summary_item)
        
        # 5-b. Post-run backup (data.js 갱신 후 스냅샷)
        if not DRY_RUN:
            auto_backup(selected_files, phase="post")

        # Call GitHub Push conditionally
        if not SKIP_GIT:
            commit_and_push_to_github()
        else:
            print("  [Info] --no-git 플래그로 인해 자동 Git Push를 건너뜁니다.")
        
    except Exception as e:
        print(f"\n❌ 파이프라인 실행 중 오류 발생: {e}")
        import traceback
        traceback.print_exc()
    finally:
        if excel:
            try:
                excel.ScreenUpdating = True
                excel.DisplayAlerts = True
                excel.Quit()
                print("\nExcel COM 연결 해제 완료.")
            except:
                pass
                
    print("\n" + "=" * 70)
    print("자동화 갱신 작업 완료! index.html을 열어 대시보드를 새로고침하세요.")
    print("=" * 70)

if __name__ == "__main__":
    main()
