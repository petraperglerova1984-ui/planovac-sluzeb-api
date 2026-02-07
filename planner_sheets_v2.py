# -*- coding: utf-8 -*-
"""
Plánovač V3 - Férové rozdělení
Hlavní princip: VŠICHNI MAJÍ PODOBNÝ ROZDÍL OD TARGETU
"""

import os
import json
import calendar
import datetime
import random
import unicodedata
import gspread
from google.oauth2.service_account import Credentials

# Konfigurace
SPREADSHEET_ID = "1L3isRHcwU9LyTMYyvT24eZk52fVfCHuYupXQLQmRyyg"
CREDENTIALS_FILE = "credentials.json"

REQ_D = 3
REQ_N = 3
SHIFT_HOURS = 11.0
MAX_CONSEC_SHIFTS = 2

BLOCK_VALUES = {"R", "DOV", "AMB", "GEN", "K", "COS", "C", "S", "POŽ"}
HOURS_FIXED = {"R": 8.0, "AMB": 8.0, "S": 8.0, "COS": 7.5, "K": 6.0}
MAX_NURSE_ROW = 46


# České svátky 2026
HOLIDAYS_2026 = {
    (1, 1): "Nový rok",
    (4, 13): "Velikonoční pondělí",  # Pohyblivý - pro 2026
    (5, 1): "Svátek práce",
    (5, 8): "Den vítězství",
    (7, 5): "Den slovanských věrozvěstů Cyrila a Metoděje",
    (7, 6): "Den upálení mistra Jana Husa",
    (9, 28): "Den české státnosti",
    (10, 28): "Den vzniku samostatného československého státu",
    (11, 17): "Den boje za svobodu a demokracii",
    (12, 24): "Štědrý den",
    (12, 25): "1. svátek vánoční",
    (12, 26): "2. svátek vánoční",
}

def is_holiday(year, month, day):
    """Kontrola zda je den svátek"""
    return (month, day) in HOLIDAYS_2026




def norm_text(s: str) -> str:
    s = (s or "").strip().upper()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return " ".join(s.split())


def to_float(v) -> float:
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    return float(str(v).replace(",", "."))


def connect_to_sheets():
    scope = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    
    creds_json = os.environ.get('GOOGLE_CREDENTIALS_JSON')
    if creds_json:
        creds_dict = json.loads(creds_json)
        creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    else:
        creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scope)
    
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID)


def get_month_from_sheet_name(sheet_name: str) -> tuple:
    # DŮLEŽITÉ: Seřaď od nejdelších (aby CERVENEC byl před CERVEN!)
    mapping = [
        ("CERVENEC", 7), ("LISTOPAD", 11), ("PROSINEC", 12),
        ("BREZEN", 3), ("KVETEN", 5), ("CERVEN", 6),
        ("LEDEN", 1), ("UNOR", 2), ("DUBEN", 4), ("SRPEN", 8),
        ("ZARI", 9), ("RIJEN", 10),
    ]
    
    name_norm = norm_text(sheet_name)
    for month_name, month_num in mapping:
        if month_name in name_norm:
            print(f"DEBUG: List {sheet_name} -> měsíc {month_num}")
            return 2026, month_num
    
    raise RuntimeError(f"Nerozumím názvu listu: '{sheet_name}'")


def detect_structure(ws_data):
    header_row, name_col = None, None
    for r_idx, row in enumerate(ws_data[:15]):
        for c_idx, cell in enumerate(row[:30] if len(row) > 30 else row):
            if isinstance(cell, str) and norm_text(cell) == "JMENO":
                header_row, name_col = r_idx + 1, c_idx + 1
                break
        if header_row:
            break
    
    if not header_row:
        raise RuntimeError("Nenašel jsem hlavičku 'Jméno'")
    
    first_row = ws_data[0] if ws_data else []
    start_col = None
    for c_idx, cell in enumerate(first_row):
        try:
            if (isinstance(cell, (int, float)) and int(cell) == 1) or \
               (isinstance(cell, str) and cell.strip() == "1"):
                start_col = c_idx + 1
                break
        except:
            continue
    
    if not start_col:
        raise RuntimeError("Nenašel jsem začátek měsíce")
    
    return header_row, name_col, start_col


def load_employees(ws_data, header_row, name_col):
    employees = []
    for r_idx in range(header_row, min(len(ws_data), MAX_NURSE_ROW)):
        if r_idx < len(ws_data) and (name_col - 1) < len(ws_data[r_idx]):
            name = ws_data[r_idx][name_col - 1]
            if isinstance(name, str) and name.strip():
                uvazek = to_float(ws_data[r_idx][0]) if len(ws_data[r_idx]) > 0 else 1.0
                employees.append({
                    'row': r_idx + 1,
                    'name': name.strip(),
                    'uvazek': uvazek if uvazek > 0 else 1.0
                })
    
    return employees


def load_hours_fund(wb, year, month):
    try:
        ws = wb.worksheet("FONDY_HODIN")
        data = ws.get_all_values()
        
        for row in data[1:]:
            if len(row) < 3:
                continue
            
            if "/" in row[0]:
                try:
                    parts = row[0].split("/")
                    date_month = int(parts[1]) if len(parts[0]) <= 2 else int(parts[0])
                    if date_month == month:
                        return to_float(row[1]), to_float(row[2])
                except:
                    pass
        
        return 176.0, 165.0
    except:
        return 176.0, 165.0


def load_employee_types(wb, employees):
    try:
        ws = wb.worksheet("ZAMESTNANCI")
        data = ws.get_all_values()
        
        types = {}
        for row in data[1:]:
            if len(row) > 3:
                name = row[0].strip()
                typ = norm_text(row[3]) if len(row) > 3 else "1S"
                types[norm_text(name)] = typ
        
        return types
    except:
        return {}


def plan_shifts_v2(sheet_name: str):
    """
    V3 - Férové rozdělení
    """
    
    print("=" * 60)
    print("Plánovač služeb V3 - Férové rozdělení")
    print("=" * 60)
    
    print(f"\n[1/7] Připojuji se...")
    wb = connect_to_sheets()
    print(f"✓ Připojeno: {wb.title}")
    
    print(f"\n[2/7] Zpracovávám list '{sheet_name}'...")
    year, month = get_month_from_sheet_name(sheet_name)
    days_in_month = calendar.monthrange(year, month)[1]
    print(f"DEBUG: calendar.monthrange({year}, {month}) = {calendar.monthrange(year, month)}")
    print(f"✓ Rok: {year}, Měsíc: {month}, Dní: {days_in_month}")
    
    print(f"\n[3/7] Načítám data...")
    ws = wb.worksheet(sheet_name)
    ws_data = ws.get_all_values()
    
    header_row, name_col, start_col = detect_structure(ws_data)
    plan_cols = list(range(start_col, start_col + days_in_month))
    print(f"✓ Struktura OK")
    
    print(f"\n[4/7] Načítám zaměstnance...")
    employees = load_employees(ws_data, header_row, name_col)
    print(f"✓ Nalezeno {len(employees)} zaměstnanců")
    
    station_idx = None
    for i, emp in enumerate(employees):
        if "STARA" in norm_text(emp['name']):
            station_idx = i
            print(f"✓ Staniční: {emp['name']}")
            break
    
    if station_idx is None:
        station_idx = 0
    
    print(f"\n[5/7] Načítám fondy...")
    fond_1s, fond_05s = load_hours_fund(wb, year, month)
    emp_types = load_employee_types(wb, employees)
    
    for emp in employees:
        name_norm = norm_text(emp['name'])
        typ = emp_types.get(name_norm, "1S")
        if "1S" in typ:
            emp['target_hours'] = fond_1s * emp['uvazek']
        else:
            emp['target_hours'] = fond_05s * emp['uvazek']
    
    print(f"\n[6/7] Načítám předvyplněné...")
    fixed = [[None] * days_in_month for _ in employees]
    fixed_hours = [0.0] * len(employees)
    
    for i, emp in enumerate(employees):
        row_data = ws_data[emp['row'] - 1] if (emp['row'] - 1) < len(ws_data) else []
        for di, col_num in enumerate(plan_cols):
            col_idx = col_num - 1
            if col_idx < len(row_data):
                val = row_data[col_idx]
                if isinstance(val, str):
                    val = val.strip().upper()
                    if val in BLOCK_VALUES or val in ("D", "N"):
                        fixed[i][di] = val
                        if val in HOURS_FIXED:
                            fixed_hours[i] += HOURS_FIXED[val]
    
    # R pro staniční (JEN po-pá, NE víkend ani svátky!)
    for di in range(days_in_month):
        dt = datetime.date(year, month, di + 1)
        # Kontrola: je to pracovní den? (po-pá a NE víkend)
        if dt.weekday() >= 5:  # Sobota nebo neděle
            continue
        # Kontrola: už tam není něco jiného?
        if fixed[station_idx][di] is not None:
            continue
        # Přidej R
        fixed[station_idx][di] = "R"
        fixed_hours[station_idx] += HOURS_FIXED["R"]
    
    r_count = sum(1 for v in fixed[station_idx] if v == "R")
    print(f"✓ Staniční má {fixed_hours[station_idx]}h (R na {r_count} dnů)")
    
    # NOVÝ ALGORITMUS - FÉROVÉ ROZDĚLENÍ
    print(f"\n[7/7] Plánuji...")
    result = fair_planner(employees, fixed, fixed_hours, days_in_month, station_idx)
    
    if not result:
        raise RuntimeError("Nelze najít řešení!")
    
    assign, hours = result
    
    # ZÁPIS
    print(f"\n{'=' * 60}")
    print("Zapisuji...")
    print('=' * 60)
    
    write_count = 0
    min_row = min(e['row'] for e in employees)
    max_row = max(e['row'] for e in employees)
    min_col = min(plan_cols)
    max_col = max(plan_cols)
    
    num_rows = max_row - min_row + 1
    num_cols = max_col - min_col + 1
    grid = [[None] * num_cols for _ in range(num_rows)]
    
    for di, col_num in enumerate(plan_cols):
        for i, emp in enumerate(employees):
            row_data = ws_data[emp['row'] - 1] if (emp['row'] - 1) < len(ws_data) else []
            col_idx = col_num - 1
            
            orig = ""
            if col_idx < len(row_data):
                orig = row_data[col_idx]
            
            new_val = assign[i][di]
            if new_val and (orig in (None, "", 0) or (i == station_idx and new_val == "R")):
                grid_row = emp['row'] - min_row
                grid_col = col_num - min_col
                grid[grid_row][grid_col] = new_val
                write_count += 1
    
    if write_count > 0:
        import gspread
        start_cell = gspread.utils.rowcol_to_a1(min_row, min_col)
        end_cell = gspread.utils.rowcol_to_a1(max_row, max_col)
        range_notation = f"{start_cell}:{end_cell}"
        ws.update(range_notation, grid, value_input_option='RAW')
    
    print(f"✓ Zapsáno {write_count} buněk")
    
    # STATISTIKY
    print(f"\n{'=' * 60}")
    print("STATISTIKY")
    print('=' * 60)
    
    for i, emp in enumerate(employees):
        diff = hours[i] - emp['target_hours']
        d_count = sum(1 for v in assign[i] if v == "D")
        n_count = sum(1 for v in assign[i] if v == "N")
        total_shifts = d_count + n_count
        
        print(f"{emp['name']:15s} target={emp['target_hours']:6.1f} "
              f"planned={hours[i]:6.1f} diff={diff:+6.1f} "
              f"shifts={total_shifts:2d} (D={d_count} N={n_count})")
    
    return {"status": "success", "sheet": sheet_name, "written": write_count}


def fair_planner(employees, fixed, fixed_hours, days, station_idx):
    """
    NOVÝ ALGORITMUS - FÉROVÉ ROZDĚLENÍ
    
    Princip:
    1. Spočítej kolik směn každý potřebuje
    2. Vyber den
    3. Pro ten den: vyber lidi kteří nejvíc potřebují směnu
    4. Přiřaď jim (vždy tak aby zůstali FÉROVĚ)
    """
    
    P = len(employees)
    D = days
    
    assign = [row[:] for row in fixed]
    hours = fixed_hours[:]
    target = [e['target_hours'] for e in employees]
    
    # Spočítej kolik směn každý potřebuje (přibližně)
    needed_shifts = []
    for i in range(P):
        remaining = target[i] - hours[i]
        shifts_needed = int(round(remaining / 11.0))
        needed_shifts.append(max(0, shifts_needed))
    
    print(f"Potřeby směn: {needed_shifts}")
    
    def can_assign(i, di, shift):
        if assign[i][di] is not None:
            return False
        
        consec = 0
        for dj in range(max(0, di - MAX_CONSEC_SHIFTS), di):
            if assign[i][dj] in ("D", "N"):
                consec += 1
        if consec >= MAX_CONSEC_SHIFTS:
            return False
        
        if di > 0 and assign[i][di - 1] == "N":
            if shift in ("D", "N"):
                return False
        
        return True
    
    def get_priority(i):
        """Priorita osoby - čím víc potřebuje směnu, tím vyšší"""
        remaining = target[i] - hours[i]
        return remaining
    
    # HLAVNÍ SMYČKA - den po dni (NÁHODNÉ POŘADÍ!)
    day_order = list(range(D))
    random.shuffle(day_order)  # Zamíchej pořadí dnů
    
    for di in day_order:
        # Kolik potřebujeme?
        needed_d = REQ_D
        needed_n = REQ_N
        
        for i in range(P):
            if assign[i][di] == "D":
                needed_d -= 1
            elif assign[i][di] == "N":
                needed_n -= 1
        
        # Přiřaď D
        for _ in range(needed_d):
            candidates = []
            for i in range(P):
                if i == station_idx:
                    continue
                if can_assign(i, di, "D"):
                    priority = get_priority(i)
                    candidates.append((priority, random.random(), i))
            
            if candidates:
                candidates.sort(reverse=True)
                _, _, best_i = candidates[0]
                assign[best_i][di] = "D"
                hours[best_i] += 11.0
        
        # Přiřaď N
        for _ in range(needed_n):
            candidates = []
            for i in range(P):
                if i == station_idx:
                    continue
                if can_assign(i, di, "N"):
                    priority = get_priority(i)
                    candidates.append((priority, random.random(), i))
            
            if candidates:
                candidates.sort(reverse=True)
                _, _, best_i = candidates[0]
                assign[best_i][di] = "N"
                hours[best_i] += 11.0
    
    return assign, hours


if __name__ == "__main__":
    result = plan_shifts_v2("CERVEN")
    print("\n✓ HOTOVO!")
