# ============================================================
#  CONFIG  –  change everything here
# ============================================================
import os
from datetime import datetime
from dateutil.relativedelta import relativedelta

BASE_DIR = r"C:\Users\khchu1\Desktop\Brian2\EPD_testing_delete"

# ------------------------------------------------------------
# 1. EXPORT
# ------------------------------------------------------------
EXPORT = {
    "source_file"   : "ExportsT.xlsx",
    "dest_pattern"  : r"^Exports of recyclables by destination\((\d{6})\)\.xlsx$",
    "sheets": {
        # main DX sheet
        "DX by destination": {
            "src_sheet"     : "DX - Result",
            "month_col"     : "D5:D93",          # dynamic column (Row 4 logic)
            "ytd_range"     : "E5:F93",          # fixed → CM5
            "ytd_dest"      : "CM5",
            "title_template": "Table 10  Trade volume (Thousand tonnes) of domestic exports of selected recyclables by destination, 2016 to {month_year}",
        },
        # RX / TX sheets (same structure)
        "RX by destination": {
            "src_sheet" : "RX - Result",
            "month_col" : "D5:D93",
            "ytd_range" : "E5:F93",
            "ytd_dest"  : "CM5",
        },
        "DXRX by destination ": {   # ← add the trailing space to fit the file formating stored in the drive
            "src_sheet" : "TX - Result",
            "month_col" : "D5:D93",
            "ytd_range" : "E5:F93",
            "ytd_dest"  : "CM5",
        },
    },
    # Top-10 special sheet
    "top10": {
        "src_sheet"       : "TX - top10",
        "src_range"       : "A3:C62",
        "dest_name_pattern": r"DXRX - .* \(top 10\)",   # will be renamed
        "paste_start"     : "B5",
    },
}

# ------------------------------------------------------------
# 2. IMPORT
# ------------------------------------------------------------
IMPORT = {
    "source_file"  : "ImportsT.xlsx",
    "dest_pattern" : "Imports of recyclables by country(*).xlsx",
    "sheets": [
        {
            "src"          : "CC - Result",
            "dest"         : "IM by CC",
            "month_range"  : "D5:D123",
            "ytd_range"    : "E5:F123",
            "ytd_dest"     : "CM5:CN123",
            "a1_replace"   : True,          # replace month abbr in A1
        },
        {
            "src"          : "CO - Result",
            "dest"         : "IM by CO",
            "month_range"  : "D5:D123",
            "ytd_range"    : "E5:F123",
            "ytd_dest"     : "CM5:CN123",
            "a1_replace"   : True,
        },
    ],
}

# ------------------------------------------------------------
# 3. TRADE
# ------------------------------------------------------------
TRADE = {
    "source_file"  : "TradeT_C.xlsx",
    "dest_pattern" : "Trade data by type and year(*).xlsx",
    "sheets"       : ["DX", "IM", "RX"],
    "src_sheets"   : ["DX - Result", "IM - Result", "RX - Result"],
    "header_range" : "C3:AP4",
    "header_dest"  : "C18",
    "month_src_lookup": "B18:B29",   # ← NEW: where month numbers 1-12 live in SOURCE
    "month_lookup" : "B259:B270",    # destination still uses month abbreviations
    "hide_row"     : 18,
}
# ------------------------------------------------------------
# 4. VOLUME
# ------------------------------------------------------------
VOLUME = {
    "source_file"  : "VolumeT.xlsx",
    "dest_pattern" : "Volume and unit values of recyclables(*).xlsx",
    "jobs": [
        # (src_sheet, src_range, dest_sheet, dest_range)
        ("DX - Result", "D5:I54", "DX", "AI5:AN54"),
        ("IM - Result", "D5:I58", "IM", "AI5:AN58"),
        ("RX - Result", "D5:I58", "RX", "AI5:AN58"),
    ],
}

# ============================================================
#  SHARED HELPERS
# ============================================================
import re
import glob
import win32com.client as win32

xlPasteValues  = -4163
xlPasteFormats = -4122
xlValues       = -4163

import pythoncom
import gc

def get_excel():
    # Kill any leftover Excel processes first (optional but very effective)
    try:
        os.system('taskkill /F /IM EXCEL.EXE >nul 2>&1')
    except:
        pass

    pythoncom.CoInitialize()
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    excel.EnableEvents = False
    return excel

def find_dest_file(pattern, is_regex=False):
    if is_regex:
        for f in os.listdir(BASE_DIR):
            m = re.match(pattern, f)
            if m:
                return os.path.join(BASE_DIR, f), m.group(1)
        raise FileNotFoundError(f"No file matching regex: {pattern}")
    else:
        files = glob.glob(os.path.join(BASE_DIR, pattern))
        if not files:
            raise FileNotFoundError(f"No file matching: {pattern}")
        path = files[0]
        m = re.search(r"\((\d{6})\)", os.path.basename(path))
        yyyymm = m.group(1) if m else None
        return path, yyyymm

def paste_values_skip_formulas(ws_src, src_addr, ws_dest, dest_start):
    """Paste values, skipping any destination cell that already has a formula."""
    src_range = ws_src.Range(src_addr)
    vals = src_range.Value
    rows = src_range.Rows.Count
    cols = src_range.Columns.Count
    start = ws_dest.Range(dest_start)
    sr, sc = start.Row, start.Column

    for r in range(rows):
        for c in range(cols):
            cell = ws_dest.Cells(sr + r, sc + c)
            if cell.HasFormula:
                continue
            if rows == 1 and cols == 1:
                cell.Value = vals
            else:
                cell.Value = vals[r][c] if vals else None

def get_next_month_info(yyyymm):
    d = datetime.strptime(yyyymm, "%Y%m")
    nxt = d + relativedelta(months=1)
    return {
        "curr_abbr": d.strftime("%b"),
        "next_abbr": nxt.strftime("%b"),
        "next_year": nxt.strftime("%Y"),
        "search_str": f"{yyyymm[:4]}/{yyyymm[4:]}",
        "target_text": f"{nxt.strftime('%b')} {nxt.strftime('%Y')}",
    }

# ============================================================
#  1. EXPORT
# ============================================================
def run_export():
    print("\n========== EXPORT ==========")
    src_path = os.path.join(BASE_DIR, EXPORT["source_file"])
    dest_path, yyyymm = find_dest_file(EXPORT["dest_pattern"], is_regex=True)
    info = get_next_month_info(yyyymm)
    print(f"Source : {src_path}")
    print(f"Dest   : {dest_path}  ({info['search_str']})")

    excel = get_excel()
    try:
        wb_src  = excel.Workbooks.Open(src_path, UpdateLinks=False, ReadOnly=True)
        wb_dest = excel.Workbooks.Open(dest_path, UpdateLinks=False)
        print("Sheets in dest:", [sh.Name for sh in wb_dest.Sheets])
        # ---- normal sheets (DX / RX / DXRX) ----
        for dest_name, cfg in EXPORT["sheets"].items():
            print(f"\n--- {cfg['src_sheet']} → {dest_name} ---")
            ws_src  = wb_src.Sheets(cfg["src_sheet"])
            ws_dest = wb_dest.Sheets(dest_name)

            # dynamic month column (Row 4 logic)
            found = ws_dest.Rows(4).Find(What=info["search_str"], LookIn=xlValues)
            if not found:
                raise ValueError(f"'{info['search_str']}' not found in Row 4 of {dest_name}")
            target_col = found.Column + 1
            print(f"  Header at col {found.Column} → paste into col {target_col}")

            ws_src.Range(cfg["month_col"]).Copy()
            ws_dest.Cells(5, target_col).PasteSpecial(Paste=xlPasteValues)

            # format brush from left column
            last_row = 93
            fmt_src = ws_dest.Range(ws_dest.Cells(5, found.Column),
                                    ws_dest.Cells(last_row, found.Column))
            fmt_dst = ws_dest.Range(ws_dest.Cells(5, target_col),
                                    ws_dest.Cells(last_row, target_col))
            fmt_src.Copy()
            fmt_dst.PasteSpecial(Paste=xlPasteFormats)
            excel.CutCopyMode = False

            # fixed YTD / YoY
            ws_src.Range(cfg["ytd_range"]).Copy()
            ws_dest.Range(cfg["ytd_dest"]).PasteSpecial(Paste=xlPasteValues)
            excel.CutCopyMode = False
            print("  Month + YTD pasted.")

            # title update (only for the main DX sheet)
            if "title_template" in cfg:
                new_title = cfg["title_template"].format(month_year=info["target_text"])
                ws_dest.Range("A1").Value = new_title
                print(f"  A1 updated → {new_title}")

        # ---- Top-10 sheet ----
        print("\n--- Top-10 sheet ---")
        tcfg = EXPORT["top10"]
        ws_src = wb_src.Sheets(tcfg["src_sheet"])

        # collect candidates
        candidates = [sh for sh in wb_dest.Sheets
                    if re.search(tcfg["dest_name_pattern"], sh.Name)]

        has_curr = any(info["curr_abbr"] in sh.Name for sh in candidates)
        has_next = any(info["next_abbr"] in sh.Name for sh in candidates)

        # Re-run detection: sheet already renamed → stop
        if not has_curr and has_next:
            print("\n*** RE-RUN DETECTED ***")
            print(f"Top-10 sheet is already named with next month ({info['next_abbr']}).")
            print("This destination file has already been updated.")
            print("Please restore it from backup before running again.")
            print("Stopping EXPORT.")
            return False          # ← important

        # normal first-run: prefer sheet that still has current month
        ws_top = None
        for sh in candidates:
            if info["curr_abbr"] in sh.Name:
                ws_top = sh
                break

        # fallback
        if not ws_top:
            for sh in candidates:
                if "2025" not in sh.Name:
                    ws_top = sh
                    break

        if not ws_top:
            raise ValueError(f"Top-10 sheet not found. Candidates: {[s.Name for s in candidates]}")

        # filter logic (same as original)
        source_vals = ws_src.Range(tcfg["src_range"]).Value
        filtered = []
        section_headers = ["ferrous metals", "non-ferrous metals", "total metals", "plastics", "paper"]
        current_cat = ""
        for row in source_vals:
            col_a = str(row[0] or "").strip().lower()
            col_b = str(row[1] or "").strip()
            col_c = row[2]
            if col_a in section_headers:
                current_cat = col_a
            elif col_b.lower() in section_headers:
                current_cat = col_b.lower()

            is_zero = False
            try:
                is_zero = abs(float(col_c or 0)) < 1e-9
            except:
                pass

            skip = (col_b.lower() == "usa" and is_zero) or \
                   (current_cat == "paper" and col_b.lower() == "other economies" and is_zero)
            if not skip:
                filtered.append([row[1], row[2]])

        # clear & paste
        ws_top.Range("B5:C65").ClearContents()
        end_row = 5 + len(filtered) - 1
        dest_rng = ws_top.Range(f"B5:C{end_row}")
        dest_rng.Value = [tuple(r) for r in filtered]
        dest_rng.Copy()
        dest_rng.PasteSpecial(Paste=xlPasteValues)
        excel.CutCopyMode = False

        # replace month text + rename sheet
        ws_top.UsedRange.Replace(What=info["curr_abbr"], Replacement=info["next_abbr"],
                                 LookAt=2, SearchOrder=1, MatchCase=True)
        new_name = f"DXRX - {info['next_abbr']} {info['next_year']} (top 10)"
        ws_top.Name = new_name
        print(f"  Top-10 updated and renamed to '{new_name}'")

        wb_dest.Save()
        print("\nEXPORT finished successfully.")
        print("\nEXPORT finished successfully.")
        return True               # ← add this
    finally:
        excel.ScreenUpdating = True
        excel.EnableEvents = True
        try: wb_src.Close(False)
        except: pass
        try: wb_dest.Close(False)
        except: pass
        excel.Quit()

# ============================================================
#  2. IMPORT
# ============================================================
def run_import():
    print("\n========== IMPORT ==========")
    src_path = os.path.join(BASE_DIR, IMPORT["source_file"])
    dest_path, yyyymm = find_dest_file(IMPORT["dest_pattern"])
    info = get_next_month_info(yyyymm)
    print(f"Source : {src_path}")
    print(f"Dest   : {dest_path}")

    excel = get_excel()
    try:
        wb_src  = excel.Workbooks.Open(src_path, UpdateLinks=False, ReadOnly=True)
        wb_dest = excel.Workbooks.Open(dest_path, UpdateLinks=False)

        for job in IMPORT["sheets"]:
            print(f"\n--- {job['src']} → {job['dest']} ---")
            ws_src  = wb_src.Sheets(job["src"])
            ws_dest = wb_dest.Sheets(job["dest"])

            # A1 month replace
            if job.get("a1_replace"):
                old = str(ws_dest.Range("A1").Value or "")
                new = old.replace(info["curr_abbr"], info["next_abbr"])
                ws_dest.Range("A1").Value = new
                print(f"  A1: '{old}' → '{new}'")

            # dynamic month column
            found = ws_dest.Rows(4).Find(What=info["search_str"], LookIn=xlValues)
            if not found:
                raise ValueError(f"'{info['search_str']}' not found in Row 4")
            target_col = found.Column + 1
            print(f"  Header at col {found.Column} → paste into col {target_col}")

            # month values + format from left
            src_m = ws_src.Range(job["month_range"])
            dst_m = ws_dest.Range(ws_dest.Cells(5, target_col),
                                  ws_dest.Cells(123, target_col))
            src_m.Copy()
            dst_m.PasteSpecial(Paste=xlPasteValues)
            excel.CutCopyMode = False

            fmt = ws_dest.Range(ws_dest.Cells(5, found.Column),
                                ws_dest.Cells(123, found.Column))
            fmt.Copy()
            dst_m.PasteSpecial(Paste=xlPasteFormats)
            excel.CutCopyMode = False

            # YTD / YoY (values only)
            ws_src.Range(job["ytd_range"]).Copy()
            ws_dest.Range(job["ytd_dest"]).PasteSpecial(Paste=xlPasteValues)
            excel.CutCopyMode = False
            print("  Month + YTD pasted.")

        wb_dest.Save()
        print("\nIMPORT finished successfully.")
    finally:
        excel.ScreenUpdating = True
        excel.EnableEvents = True
        try: wb_src.Close(False)
        except: pass
        try: wb_dest.Close(False)
        except: pass
        excel.Quit()

# ============================================================
#  3. TRADE
# ============================================================
def run_trade():
    print("\n========== TRADE ==========")
    src_path = os.path.join(BASE_DIR, TRADE["source_file"])
    dest_path, yyyymm = find_dest_file(TRADE["dest_pattern"])
    info = get_next_month_info(yyyymm)
    print(f"Source : {src_path}")
    print(f"Dest   : {dest_path}  (paste into {info['next_abbr']})")

    excel = get_excel()
    try:
        wb_src  = excel.Workbooks.Open(src_path, UpdateLinks=False, ReadOnly=True)
        wb_dest = excel.Workbooks.Open(dest_path, UpdateLinks=False)

        for src_name, dest_name in zip(TRADE["src_sheets"], TRADE["sheets"]):
            print(f"\n--- {src_name} → {dest_name} ---")
            ws_src  = wb_src.Sheets(src_name)
            ws_dest = wb_dest.Sheets(dest_name)

            # 1) header rows (skip formulas)
            print("  Pasting header rows (skip formulas)...")
            paste_values_skip_formulas(ws_src, TRADE["header_range"],
                                    ws_dest, TRADE["header_dest"])

            # 2) find SOURCE row for the NEXT month
            next_month_num = (datetime.strptime(yyyymm, "%Y%m") + relativedelta(months=1)).month

            src_month_cell = None
            for cell in ws_src.Range(TRADE["month_src_lookup"]):
                try:
                    if int(cell.Value) == next_month_num:
                        src_month_cell = cell
                        break
                except (TypeError, ValueError):
                    continue

            if not src_month_cell:
                raise ValueError(f"Month number {next_month_num} not found in source {TRADE['month_src_lookup']}")

            src_row = src_month_cell.Row
            month_src_range = f"C{src_row}:AP{src_row}"
            print(f"  Source month {next_month_num} found at row {src_row} → {month_src_range}")

            # 3) locate destination paste row
            month_cell = None
            for cell in ws_dest.Range(TRADE["month_lookup"]):
                val = str(cell.Value or "").strip()
                if val.lower() == info["curr_abbr"].lower():
                    month_cell = cell
                    break
            if not month_cell:
                raise ValueError(f"'{info['curr_abbr']}' not found in {TRADE['month_lookup']}")
            target_row = month_cell.Row + 1
            print(f"  Found {info['curr_abbr']} at B{month_cell.Row} → paste to row {target_row} ({info['next_abbr']})")

            paste_values_skip_formulas(ws_src, month_src_range,
                                    ws_dest, f"C{target_row}")

            # 4) hide row
            ws_dest.Rows(TRADE["hide_row"]).Hidden = True
            print(f"  Row {TRADE['hide_row']} hidden.")

        wb_dest.Save()
        print("\nTRADE finished successfully.")
    finally:
        excel.ScreenUpdating = True
        excel.EnableEvents = True
        try: wb_src.Close(False)
        except: pass
        try: wb_dest.Close(False)
        except: pass
        excel.Quit()

# ============================================================
#  4. VOLUME
# ============================================================
def run_volume():
    print("\n========== VOLUME ==========")
    src_path = os.path.join(BASE_DIR, VOLUME["source_file"])
    dest_path, yyyymm = find_dest_file(VOLUME["dest_pattern"])
    info = get_next_month_info(yyyymm)
    print(f"Source : {src_path}")
    print(f"Dest   : {dest_path}")

    excel = get_excel()
    try:
        wb_src  = excel.Workbooks.Open(src_path, UpdateLinks=False, ReadOnly=True)
        wb_dest = excel.Workbooks.Open(dest_path, UpdateLinks=False)


        for src_sheet, src_rng, dest_sheet, dest_rng in VOLUME["jobs"]:
            print(f"\n--- {src_sheet} → {dest_sheet} ---")
            print(f"  {src_rng} → {dest_rng}")
            ws_src  = wb_src.Sheets(src_sheet)
            ws_dest = wb_dest.Sheets(dest_sheet)

            # Update cell A1 month for volumn
            old = str(ws_dest.Range("A1").Value or "")
            new = old.replace(info["curr_abbr"], info["next_abbr"])
            if old != new:
                ws_dest.Range("A1").Value = new
                print(f"  A1: '{old}' → '{new}'")
            else:
                print(f"  A1 unchanged: '{old}'")


            ws_src.Range(src_rng).Copy()
            ws_dest.Range(dest_rng).PasteSpecial(Paste=xlPasteValues)
            excel.CutCopyMode = False
            print("  Done.")

        wb_dest.Save()
        print("\nVOLUME finished successfully.")
    finally:
        excel.ScreenUpdating = True
        excel.EnableEvents = True
        try: wb_src.Close(False)
        except: pass
        try: wb_dest.Close(False)
        except: pass
        excel.Quit()


def rename_all_dest_files():
    print("\n========== RENAME DESTINATION FILES ==========")
    jobs = [
        (EXPORT["dest_pattern"], True),    # regex
        (IMPORT["dest_pattern"], False),
        (TRADE["dest_pattern"],  False),
        (VOLUME["dest_pattern"], False),
    ]
    for pattern, is_regex in jobs:
        try:
            path, yyyymm = find_dest_file(pattern, is_regex=is_regex)
            if not yyyymm:
                print(f"  Skip (no date in name): {os.path.basename(path)}")
                continue

            nxt = datetime.strptime(yyyymm, "%Y%m") + relativedelta(months=1)
            new_yyyymm = nxt.strftime("%Y%m")

            old_name = os.path.basename(path)
            new_name = old_name.replace(f"({yyyymm})", f"({new_yyyymm})")
            new_path = os.path.join(os.path.dirname(path), new_name)

            if os.path.exists(new_path):
                print(f"  SKIP (already exists): {new_name}")
                continue

            os.rename(path, new_path)
            print(f"  {old_name}")
            print(f"    → {new_name}")
        except FileNotFoundError:
            print(f"  Not found: {pattern}")
        except Exception as e:
            print(f"  Error: {pattern} → {e}")

# ============================================================
#  RUN EVERYTHING (or call individually)
# ============================================================


if __name__ == "__main__":
    export_ok = run_export()

    if export_ok is False:
        print("\nProgram stopped because of re-run detection. No further steps, no rename.")
    else:
        run_import()
        run_trade()
        run_volume()
        rename_all_dest_files()
        pass