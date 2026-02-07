# -*- coding: utf-8 -*-
"""
Google Sheets adapter pro planner
Čte data z Google Sheets, spustí plánování a zapíše výsledky zpět
"""

import os
import json
import calendar
import datetime
import math
import random
import unicodedata
from typing import List, Tuple, Dict, Any
import gspread
from google.oauth2.service_account import Credentials

# Konfigurace
SPREADSHEET_ID = "1L3isRHcwU9LyTMYyvT24eZk52fVfCHuYupXQLQmRyyg"  # AKTUALIZOVANÉ ID!
CREDENTIALS_FILE = "credentials.json"

SHEET_ZAM = "ZAMESTNANCI"
SHEET_FONDY = "FONDY_HODIN"
SHEET_SETTINGS = "NASTAVENI"

# Plánovací parametry (stejné jako v původním kódu)
REQ_D = 3
REQ_N = 3
SHIFT_HOURS = 11.0

FORBID_N_AFTER_N = True
FORBID_D_AFTER_N = True
FORBID_SHIFT_ON_SECOND_DAY_AFTER_N = True
FORBID_N_BEFORE_ANY_PREFILL = True
MAX_CONSEC_SHIFTS = 2
MIN_FREE_WEEKENDS = 2

BLOCK_VALUES = {"R", "DOV", "AMB", "GEN", "K", "COS", "C", "S", "POŽ"}
HOURS_FIXED = {"R": 8.0, "AMB": 8.0, "S": 8.0, "COS": 7.5, "K": 6.0}

JITTER = 0.02
BEHIND_W = 1.1
OVERDUE_W = 1.3
MIX_W = 1.6

SECOND_DAY_AFTER_N_PENALTY = 2.2
N_AFTER_D_BONUS = 1.4

FREE_WEEKENDS_GOAL = 2
WEEKEND_FAIR_W = 4.0
WEEKEND_OVER_W = 10.0
FRIDAY_N_BASE_W = 2.5
FRIDAY_N_SCALE_W = 1.2

MAX_NURSE_ROW = 46


def norm_text(s: str) -> str:
    s = (s or "").strip().upper()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = " ".join(s.split())
    return s


def to_float(v) -> float:
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    return float(str(v).replace(",", "."))


def month_num_from_name(name: str) -> int:
    m = norm_text(name)
    mapping = {
        "LEDEN": 1,
        "UNOR": 2, "ÚNOR": 2,
        "BREZEN": 3, "BŘEZEN": 3,
        "DUBEN": 4,
        "KVETEN": 5, "KVĚTEN": 5,
        "CERVEN": 6, "ČERVEN": 6,
        "CERVENEC": 7, "ČERVENEC": 7,
        "SRPEN": 8,
        "ZARI": 9, "ZÁŘÍ": 9,
        "RIJEN": 10, "ŘÍJEN": 10,
        "LISTOPAD": 11,
        "PROSINEC": 12,
    }
    if m not in mapping:
        raise RuntimeError(f"Nerozumím měsíci v NASTAVENI!B1: '{name}'")
    return mapping[m]


def sheet_name_from_month_num(month_num: int) -> str:
    return {
        1: "LEDEN", 2: "UNOR", 3: "BREZEN", 4: "DUBEN",
        5: "KVETEN", 6: "CERVEN", 7: "CERVENEC", 8: "SRPEN",
        9: "ZARI", 10: "RIJEN", 11: "LISTOPAD", 12: "PROSINEC",
    }[month_num]


def prev_year_month(year: int, month: int):
    if month == 1:
        return year - 1, 12
    return year, month - 1


def connect_to_sheets():
    """Připojí se k Google Sheets"""
    scope = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    
    # Načti credentials z environment variable nebo souboru
    creds_json = os.environ.get('GOOGLE_CREDENTIALS_JSON')
    
    if creds_json:
        # Na serveru - použij environment variable
        creds_dict = json.loads(creds_json)
        creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    else:
        # Lokálně - použij soubor
        creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scope)
    
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID)


def detect_header(ws_values):
    """Najde hlavičku Jméno"""
    for r_idx, row in enumerate(ws_values[:15]):
        for c_idx, cell in enumerate(row[:30] if len(row) > 30 else row):
            if isinstance(cell, str) and norm_text(cell) == "JMENO":
                return r_idx + 1, c_idx + 1  # 1-indexed
    raise RuntimeError("Nenašel jsem hlavičku 'Jméno'.")


def detect_plan_cols(ws_values, year: int, month: int):
    """Najde plánovací sloupce (první řádek s čísly dnů)"""
    days_in_month = calendar.monthrange(year, month)[1]
    
    if not ws_values or len(ws_values) < 1:
        raise RuntimeError("List je prázdný")
    
    first_row = ws_values[0]
    start_col = None
    
    for c_idx, cell in enumerate(first_row):
        try:
            if isinstance(cell, (int, float)) and int(cell) == 1:
                start_col = c_idx + 1  # 1-indexed
                break
            elif isinstance(cell, str) and cell.strip() == "1":
                start_col = c_idx + 1
                break
        except (ValueError, TypeError):
            continue
    
    if start_col is None:
        raise RuntimeError("V řádku 1 jsem nenašel číslo '1' (začátek měsíce).")
    
    end_col = start_col + days_in_month - 1
    
    return list(range(start_col, end_col + 1)), start_col, end_col


def plan_shifts_from_sheets():
    """
    Hlavní funkce - načte data z Google Sheets, spustí plánování a zapíše výsledky
    """
    
    print("=" * 60)
    print("Planner V20 - Google Sheets verze")
    print("=" * 60)
    
    # Připoj se k tabulce
    print("\n[1/6] Připojuji se k Google Sheets...")
    wb = connect_to_sheets()
    print(f"✓ Připojeno k: {wb.title}")
    
    # Načti nastavení
    print("\n[2/6] Načítám nastavení...")
    ws_settings = wb.worksheet(SHEET_SETTINGS)
    settings_data = ws_settings.get_all_values()
    
    # Rok a měsíc z NASTAVENI (A1, B1)
    year = int(settings_data[0][0]) if settings_data and settings_data[0] else 2026
    month_name = settings_data[0][1] if settings_data and len(settings_data[0]) > 1 else "LEDEN"
    month = month_num_from_name(month_name)
    print(f"DEBUG: Hledám měsíc: {month_name} -> normalizováno: {norm_text(month_name)}")
    
    print(f"✓ Rok: {year}, Měsíc: {month_name} ({month})")
    
    # Název listu pro plánování
    sheet_name = sheet_name_from_month_num(month)
    
    # Načti plánovací list
    print(f"\n[3/6] Načítám list '{sheet_name}'...")
    ws_plan = wb.worksheet(sheet_name)
    plan_data = ws_plan.get_all_values()
    
    # Detekuj strukturu
    header_row, name_col = detect_header(plan_data)
    plan_cols, start_col, end_col = detect_plan_cols(plan_data, year, month)
    D = len(plan_cols)
    
    print(f"✓ Hlavička v řádku {header_row}, sloupec jmen {name_col}")
    print(f"✓ Plánovací sloupce: {start_col}–{end_col} ({D} dní)")
    
    # Načti seznam lidí
    print("\n[4/6] Načítám seznam zaměstnanců...")
    people_rows = []
    names = []
    
    for r_idx in range(header_row, min(len(plan_data), MAX_NURSE_ROW)):
        if r_idx < len(plan_data) and (name_col - 1) < len(plan_data[r_idx]):
            name = plan_data[r_idx][name_col - 1]
            if isinstance(name, str) and name.strip():
                people_rows.append(r_idx + 1)  # 1-indexed
                names.append(name.strip())
    
    P = len(people_rows)
    print(f"✓ Nalezeno {P} zaměstnanců")
    
    if P == 0:
        raise RuntimeError("Nenašel jsem žádné zaměstnance!")
    
    # Načti úvazky (sloupec A)
    print("\n[5/6] Načítám úvazky a fondy hodin...")
    uvazky = []
    for r_idx, row_num in enumerate(people_rows):
        row_data = plan_data[row_num - 1] if (row_num - 1) < len(plan_data) else []
        uvazek = to_float(row_data[0]) if len(row_data) > 0 else 1.0
        uvazky.append(uvazek if uvazek > 0 else 1.0)
    
    # Načti fondy hodin
    ws_fondy = wb.worksheet(SHEET_FONDY)
    fondy_data = ws_fondy.get_all_values()
    
    # Najdi řádek pro tento měsíc
    fond_1s = 0.0
    fond_0s = 0.0
    
    for row in fondy_data[1:]:  # Přeskoč hlavičku
        print(f"DEBUG FONDY: Celkem {len(fondy_data)} řádků v FONDY_HODIN")
        if len(row) > 0:
            row_month = norm_text(row[0])
            print(f"  Řádek: {row[0]} -> normalizováno: {row_month}")
            month_normalized = norm_text(month_name)
            # Zkus najít měsíc - buď podle textu nebo podle čísla v datu
            month_matched = False
            if row_month == month_normalized or month_normalized in row_month or row_month in month_normalized:
                month_matched = True
            # Zkus také podle čísla měsíce v datu (např. "01/06/2026" pro červen = měsíc 6)
            elif "/" in row[0]:
                try:
                    # Pokus se extrahovat měsíc z data
                    parts = row[0].split("/")
                    if len(parts) >= 2:
                        date_month = int(parts[1]) if len(parts[0]) <= 2 else int(parts[0])
                        if date_month == month:
                            month_matched = True
                except:
                    pass
            
            if month_matched:
                fond_1s = to_float(row[1]) if len(row) > 1 else 0.0
                print(f"  ✓ NAŠEL! fond_1s={fond_1s}, fond_0s={fond_0s}")
                fond_0s = to_float(row[2]) if len(row) > 2 else 0.0
                break
    
    print(f"✓ Fond 1S: {fond_1s}, Fond 0,5S: {fond_0s}")
    
    # Načti typy úvazků (ze ZAM, sloupec 4)
    ws_zam = wb.worksheet(SHEET_ZAM)
    zam_data = ws_zam.get_all_values()
    
    target_hours = []
    for i, name in enumerate(names):
        # Najdi zaměstnance v ZAM listu
        found = False
        for zam_row in zam_data[1:]:
            if len(zam_row) > 0 and norm_text(zam_row[0]) == norm_text(name):
                typ_uvazku = norm_text(zam_row[3]) if len(zam_row) > 3 else "1S"
                if "1S" in typ_uvazku or typ_uvazku == "1S":
                    target_hours.append(fond_1s * uvazky[i])
                else:
                    target_hours.append(fond_0s * uvazky[i])
                found = True
                break
        
        if not found:
            target_hours.append(fond_1s * uvazky[i])
    
    # Načti stávající data (fixed values)
    print("\n[6/6] Načítám předvyplněné služby...")
    fixed_shift = [["OFF"] * D for _ in range(P)]
    fixed_hours = [0.0] * P
    
    for i, row_num in enumerate(people_rows):
        row_data = plan_data[row_num - 1] if (row_num - 1) < len(plan_data) else []
        for di, col_num in enumerate(plan_cols):
            col_idx = col_num - 1
            if col_idx < len(row_data):
                cell_value = row_data[col_idx]
                if isinstance(cell_value, str):
                    cell_value = cell_value.strip().upper()
                    if cell_value in BLOCK_VALUES or cell_value in ("D", "N"):
                        fixed_shift[i][di] = cell_value
                        if cell_value in HOURS_FIXED:
                            fixed_hours[i] += HOURS_FIXED[cell_value]
    
    print(f"✓ Načteny předvyplněné služby")
    
    # Načti koncovky z předchozího měsíce
    print("\nNačítám koncovky z předchozího měsíce...")
    prev_year, prev_month = prev_year_month(year, month)
    prev_sheet_name = sheet_name_from_month_num(prev_month)
    
    prev_tail3 = [["OFF", "OFF", "OFF"] for _ in range(P)]
    
    try:
        ws_prev = wb.worksheet(prev_sheet_name)
        prev_data = ws_prev.get_all_values()
        prev_plan_cols, _, _ = detect_plan_cols(prev_data, prev_year, prev_month)
        
        # Vezmi poslední 3 dny
        last_3_cols = prev_plan_cols[-3:] if len(prev_plan_cols) >= 3 else prev_plan_cols
        
        for i, row_num in enumerate(people_rows):
            row_data = prev_data[row_num - 1] if (row_num - 1) < len(prev_data) else []
            for idx, col_num in enumerate(last_3_cols):
                col_idx = col_num - 1
                if col_idx < len(row_data):
                    cell_value = row_data[col_idx]
                    if isinstance(cell_value, str):
                        cell_value = cell_value.strip().upper()
                        if cell_value in ("D", "N"):
                            prev_tail3[i][idx] = cell_value
        
        print(f"✓ Načteny koncovky z {prev_sheet_name}")
    
    except Exception as e:
        print(f"⚠ Nelze načíst předchozí měsíc ({prev_sheet_name}): {e}")
    
    # SPUSŤ PLÁNOVÁNÍ
    print("\n" + "=" * 60)
    print("Spouštím plánovací algoritmus...")
    print("=" * 60)
    
    result = run_planning_algorithm(
        P, D, names, target_hours, fixed_shift, fixed_hours,
        prev_tail3, year, month
    )
    
    if not result:
        raise RuntimeError("Nelze sestavit rozpis (kombinace pravidel je příliš přísná).")
    
    assign, count, day_count, night_count, hours, busy_count = result
    
        # ZÁPIS VÝSLEDKŮ ZPĚT DO GOOGLE SHEETS
    print("\n" + "=" * 60)
    print("Zapisuji výsledky do Google Sheets...")
    print("=" * 60)
    
    # Připravíme všechny hodnoty najednou
    # Vytvoříme 2D mřížku s výsledky
    updates_count = 0
    
    # Najdeme rozsah který potřebujeme aktualizovat
    min_row = min(people_rows[1:]) if len(people_rows) > 1 else people_rows[0]  # Přeskočíme první (staniční)
    max_row = max(people_rows)
    min_col = min(plan_cols)
    max_col = max(plan_cols)
    
    # Vytvoříme prázdnou mřížku
    num_rows = max_row - min_row + 1
    num_cols = max_col - min_col + 1
    grid = [[None for _ in range(num_cols)] for _ in range(num_rows)]
    
    # Naplníme hodnoty které chceme zapsat
    for di, col_num in enumerate(plan_cols):
        for i, row_num in enumerate(people_rows):
            # Přeskoč staniční sestru
            if i == 0:
                continue
            
            # Přeskoč buňky, kde už něco je
            row_data = plan_data[row_num - 1] if (row_num - 1) < len(plan_data) else []
            col_idx = col_num - 1
            
            orig_value = ""
            if col_idx < len(row_data):
                orig_value = row_data[col_idx]
            
            if orig_value not in (None, "", 0):
                continue
            
            # Zapiš výsledek do mřížky
            new_value = "" if assign[i][di] == "OFF" else assign[i][di]
            
            grid_row = row_num - min_row
            grid_col = col_num - min_col
            
            if new_value:  # Jen pokud není prázdné
                grid[grid_row][grid_col] = new_value
                updates_count += 1
    
    # Zapiš celou mřížku najednou
    if updates_count > 0:
        print(f"✓ Zapisuji {updates_count} buněk najednou...")
        
        start_cell = gspread.utils.rowcol_to_a1(min_row, min_col)
        end_cell = gspread.utils.rowcol_to_a1(max_row, max_col)
        range_notation = f"{start_cell}:{end_cell}"
        
        ws_plan.update(range_notation, grid, value_input_option='RAW')
        print(f"✓ Zápis dokončen! {updates_count} buněk.")
    else:
        print("⚠ Žádné buňky k zápisu")

    # STATISTIKY
    print("\n" + "=" * 60)
    print("STATISTIKY")
    print("=" * 60)
    
    for i in range(P):
        diff = hours[i] - target_hours[i]
        print(f"{names[i]:14s} target={target_hours[i]:6.1f} planned={hours[i]:6.1f} "
              f"diff={diff:6.1f} shifts={count[i]:2d} N={night_count[i]:2d} D={day_count[i]:2d}")
    
    return {
        "year": year,
        "month": month_name,
        "people_count": P,
        "days": D,
        "statistics": [
            {
                "name": names[i],
                "target": target_hours[i],
                "planned": hours[i],
                "diff": hours[i] - target_hours[i],
                "shifts": count[i],
                "nights": night_count[i],
                "days": day_count[i]
            }
            for i in range(P)
        ]
    }


def run_planning_algorithm(P, D, names, target, fixed_shift, fixed, prev_tail3, year, month):
    """
    Zde běží původní plánovací algoritmus z planner_V7_9.py
    (backtracking s pravidly)
    """
    
    # Zjisti víkendy
    W = 0  # počet víkendů
    is_weekend = [False] * D
    weekends = []
    
    for di in range(D):
        day_num = di + 1
        dt = datetime.date(year, month, day_num)
        wd = dt.weekday()
        
        if wd in (5, 6):  # sobota nebo neděle
            is_weekend[di] = True
            
            # Přidej víkend (pátek-neděle)
            if wd == 5:  # sobota
                fri = di - 1 if di > 0 else None
                sat = di
                sun = di + 1 if di + 1 < D else None
                weekends.append((fri, sat, sun))
    
    W = len(weekends)
    
    # Staniční sestra = najdi podle jména
    station_pi = 0
    for idx, name in enumerate(names):
        if "STARA" in norm_text(name):
            station_pi = idx
            print(f"✓ Staniční sestra: {name} (index {idx})")
            break
    
    # Doplň R pro staniční sestru (po–pá v plánovacím bloku)
    for di in range(D):
        dt = datetime.date(year, month, di + 1)
        wd = dt.weekday()
        
        if 0 <= wd <= 4 and fixed_shift[station_pi][di] == "OFF":
            fixed_shift[station_pi][di] = "R"
            fixed[station_pi] += HOURS_FIXED["R"]
    
    # Inicializuj stav
    assign = [row[:] for row in fixed_shift]
    count = [0] * P
    day_count = [0] * P
    night_count = [0] * P
    hours = fixed[:]
    last_assigned = [-999] * P
    
    # Počítej již přiřazené směny z fixed
    for i in range(P):
        if i == station_pi:
            continue
        for di in range(D):
            sh = fixed_shift[i][di]
            if sh in ("D", "N"):
                count[i] += 1
                hours[i] += SHIFT_HOURS
                last_assigned[i] = di
                if sh == "N":
                    night_count[i] += 1
                else:
                    day_count[i] += 1
    
    # Víkendové struktury
    worked_weekend = [[False] * W for _ in range(P)]
    tainted_weekend = [[False] * W for _ in range(P)]
    busy_weekend = [[False] * W for _ in range(P)]
    busy_count = [0] * P
    
    # Helper funkce
    def day_in_weekend(di):
        for wi, (fri, sat, sun) in enumerate(weekends):
            if di in (fri, sat, sun):
                return wi, (fri, sat, sun)
        return None, (None, None, None)
    
    def apply_weekend_flags(i, di, sh):
        wi, (fri, sat, sun) = day_in_weekend(di)
        if wi is None:
            return []
        
        changes = []
        
        if di in (sat, sun) and sh in ("D", "N"):
            if not worked_weekend[i][wi]:
                worked_weekend[i][wi] = True
                changes.append(("worked", i, wi))
            if not busy_weekend[i][wi]:
                busy_weekend[i][wi] = True
                busy_count[i] += 1
                changes.append(("busy", i, wi))
        
        if di == fri and sh == "N":
            if not tainted_weekend[i][wi]:
                tainted_weekend[i][wi] = True
                changes.append(("tainted", i, wi))
            if not busy_weekend[i][wi]:
                busy_weekend[i][wi] = True
                busy_count[i] += 1
                changes.append(("busy", i, wi))
        
        return changes
    
    def undo_weekend_flags(i, changes):
        for change in reversed(changes):
            if change[0] == "worked":
                worked_weekend[change[1]][change[2]] = False
            elif change[0] == "busy":
                busy_weekend[change[1]][change[2]] = False
                busy_count[change[1]] -= 1
            elif change[0] == "tainted":
                tainted_weekend[change[1]][change[2]] = False
    
    def can_assign(i, di, sh):
        # Hard pravidla
        
        # Nemůže mít směnu, pokud má fixed hodnotu
        if fixed_shift[i][di] != "OFF":
            return False
        
        # Max 3 směny v kuse
        consec = 0
        for dj in range(max(0, di - MAX_CONSEC_SHIFTS), di):
            if assign[i][dj] in ("D", "N"):
                consec += 1
        if consec >= MAX_CONSEC_SHIFTS:
            return False
        
        # Po N nesmí D/N
        if di > 0 and assign[i][di - 1] == "N":
            if sh in ("D", "N") and FORBID_D_AFTER_N:
                return False
        
        # Druhý den po N taky ne
        if di > 1 and assign[i][di - 2] == "N":
            if sh in ("D", "N") and FORBID_SHIFT_ON_SECOND_DAY_AFTER_N:
                return False
        
        # Přes přelom měsíce
        if di == 0:
            if prev_tail3[i][2] == "N":
                if sh in ("D", "N"):
                    return False
        if di == 1:
            if prev_tail3[i][2] == "N":
                if sh in ("D", "N"):
                    return False
        
        # N nesmí být den před fixed hodnotou
        if sh == "N" and FORBID_N_BEFORE_ANY_PREFILL:
            if di + 1 < D and fixed_shift[i][di + 1] != "OFF":
                return False
        
        return True
    
    def score_assignment(i, di, sh):
        """Ohodnocení přiřazení (menší = lepší)"""
        sc = 0.0
        
        # Fairness: kolik už má oproti ostatním
        behind = (target[i] - hours[i]) / target[i] if target[i] > 0 else 0
        if behind > 0:
            sc -= behind * BEHIND_W
        elif behind < 0:
            sc += abs(behind) * OVERDUE_W
        
        # Mix D/N
        total_shifts = count[i] + 1
        if sh == "N":
            future_n = night_count[i] + 1
            future_d = day_count[i]
        else:
            future_n = night_count[i]
            future_d = day_count[i] + 1
        
        ratio_n = future_n / total_shifts if total_shifts > 0 else 0
        ratio_d = future_d / total_shifts if total_shifts > 0 else 0
        
        ideal_n = REQ_N / (REQ_D + REQ_N)
        ideal_d = REQ_D / (REQ_D + REQ_N)
        
        dev_n = abs(ratio_n - ideal_n)
        dev_d = abs(ratio_d - ideal_d)
        sc += (dev_n + dev_d) * MIX_W
        
        # Pattern bonus DN00
        if di > 0 and assign[i][di - 1] == "D" and sh == "N":
            sc -= N_AFTER_D_BONUS
        
        if di > 1 and assign[i][di - 2] == "N" and sh in ("D", "N"):
            sc += SECOND_DAY_AFTER_N_PENALTY
        
        # Víkendy
        wi, (fri, sat, sun) = day_in_weekend(di)
        if wi is not None:
            # Busy víkendy
            if not busy_weekend[i][wi]:
                avg_busy = sum(busy_count) / P if P > 0 else 0
                delta = busy_count[i] - avg_busy
                sc += delta * WEEKEND_FAIR_W
                
                if busy_count[i] >= W:
                    sc += WEEKEND_OVER_W
            
            # Páteční N
            if di == fri and sh == "N":
                fri_n_count = sum(1 for dj in range(D) for wj, (f2, _, _) in [day_in_weekend(dj)] 
                                  if wj is not None and dj == f2 and assign[i][dj] == "N")
                sc += FRIDAY_N_BASE_W + fri_n_count * FRIDAY_N_SCALE_W
        
        # Jitter
        sc += random.uniform(-JITTER, JITTER)
        
        return sc
    
    # Backtracking solver
    def solve_day(k):
        if k == D:
            return True
        
        di = k
        
        # Kolik D a N potřebujeme tento den?
        needed_d = REQ_D
        needed_n = REQ_N
        
        # Odečti už přiřazené
        for i in range(P):
            if assign[i][di] == "D":
                needed_d -= 1
            elif assign[i][di] == "N":
                needed_n -= 1
        
        slots = []
        if needed_d > 0:
            slots.extend(["D"] * needed_d)
        if needed_n > 0:
            slots.extend(["N"] * needed_n)
        
        if not slots:
            return solve_day(k + 1)
        
        # DFS přes sloty
        def dfs(slot_idx):
            if slot_idx == len(slots):
                return solve_day(k + 1)
            
            sh = slots[slot_idx]
            
            # Kandidáti
            candidates = []
            for i in range(P):
                if i == station_pi:
                    continue
                if assign[i][di] != "OFF":
                    continue
                if can_assign(i, di, sh):
                    sc = score_assignment(i, di, sh)
                    candidates.append((sc, i))
            
            candidates.sort()
            
            used = set()
            
            for _, i in candidates:
                if i in used:
                    continue
                
                # Zkus přiřadit
                used.add(i)
                assign[i][di] = sh
                count[i] += 1
                hours[i] += SHIFT_HOURS
                
                if sh == "N":
                    night_count[i] += 1
                else:
                    day_count[i] += 1
                
                prev_la = last_assigned[i]
                last_assigned[i] = di
                
                wk_changes = apply_weekend_flags(i, di, sh)
                
                if dfs(slot_idx + 1):
                    return True
                
                # Undo
                undo_weekend_flags(i, wk_changes)
                last_assigned[i] = prev_la
                
                if sh == "N":
                    night_count[i] -= 1
                else:
                    day_count[i] -= 1
                
                hours[i] -= SHIFT_HOURS
                count[i] -= 1
                assign[i][di] = "OFF"
                used.remove(i)
            
            return False
        
        return dfs(0)
    
    # Aplikuj víkendové flagy pro fixed
    for i in range(P):
        if i == station_pi:
            continue
        for di in range(D):
            sh = fixed_shift[i][di]
            if sh in ("D", "N"):
                apply_weekend_flags(i, di, sh)
    
    # Hledání řešení
    MAX_ATTEMPTS = 1200  # Sníženo pro rychlejší běh na serveru
    base_seed = year * 100 + month
    
    def _variance(vals):
        if not vals:
            return 0.0
        avg = sum(vals) / len(vals)
        return sum((v - avg) ** 2 for v in vals) / len(vals)
    
    def _fairness_score(_day_count, _night_count, _busy_count, _tainted_count):
        vN = _variance(_night_count)
        vW = _variance(_busy_count)
        vT = _variance(_tainted_count)
        vD = _variance(_day_count)
        
        rN = (max(_night_count) - min(_night_count)) if _night_count else 0
        rW = (max(_busy_count) - min(_busy_count)) if _busy_count else 0
        rT = (max(_tainted_count) - min(_tainted_count)) if _tainted_count else 0
        rD = (max(_day_count) - min(_day_count)) if _day_count else 0
        
        return (4.0 * vN + 2.0 * vW + 2.0 * vT + 1.0 * vD) + (3.0 * rN + 2.0 * rW + 2.0 * rT + 0.5 * rD)
    
    best = None
    best_score = None
    solved_count = 0
    
    print(f"Hledám nejlepší řešení (max {MAX_ATTEMPTS} pokusů)...")
    
    for attempt in range(1, MAX_ATTEMPTS + 1):
        if attempt % 100 == 0:
            print(f"  Pokus {attempt}/{MAX_ATTEMPTS}, nalezeno řešení: {solved_count}")
        
        random.seed(base_seed * 100000 + attempt)
        
        # Reset
        assign = [row[:] for row in fixed_shift]
        count = [0] * P
        day_count = [0] * P
        night_count = [0] * P
        hours = fixed[:]
        last_assigned = [-999] * P
        
        for i in range(P):
            if prev_tail3[i][2] in ("N", "D"):
                last_assigned[i] = -1
        
        worked_weekend = [[False] * W for _ in range(P)]
        tainted_weekend = [[False] * W for _ in range(P)]
        busy_weekend = [[False] * W for _ in range(P)]
        busy_count = [0] * P
        
        # Aplikuj fixed
        for i in range(P):
            if i == station_pi:
                continue
            for di in range(D):
                sh = fixed_shift[i][di]
                if sh not in ("D", "N"):
                    continue
                assign[i][di] = sh
                count[i] += 1
                hours[i] += SHIFT_HOURS
                last_assigned[i] = di
                if sh == "N":
                    night_count[i] += 1
                else:
                    day_count[i] += 1
        
        # Víkendové flagy pro fixed
        for i in range(P):
            if i == station_pi:
                continue
            for di in range(D):
                sh = fixed_shift[i][di]
                if sh in ("D", "N"):
                    apply_weekend_flags(i, di, sh)
        
        if not solve_day(0):
            continue
        
        solved_count += 1
        tainted_count = [sum(1 for x in tainted_weekend[i] if x) for i in range(P)]
        score = _fairness_score(day_count, night_count, busy_count, tainted_count)
        
        if best_score is None or score < best_score:
            best_score = score
            best = (assign, count, day_count, night_count, hours, busy_count)
        
        if solved_count >= 50 and attempt >= 200:
            break
    
    if best is None:
        return None
    
    print(f"\n✓ Nalezeno {solved_count} řešení, vybrán nejférovější (score={best_score:.3f})")
    
    return best


if __name__ == "__main__":
    result = plan_shifts_from_sheets()
    print("\n" + "=" * 60)
    print("HOTOVO!")
    print("=" * 60)
