# -*- coding: utf-8 -*-
"""
Plánovač služeb V2 - Nový od základu
Priorita: VYROVNANÉ HODINY
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

# Plánovací parametry
REQ_D = 3  # Denních směn denně
REQ_N = 3  # Nočních směn denně
SHIFT_HOURS = 11.0

# Hard pravidla
MAX_CONSEC_SHIFTS = 2  # Max 2 směny za sebou
FORBID_N_AFTER_N = True  # Po N nesmí N
FORBID_D_AFTER_N = True  # Po N nesmí D (musí volno)

# Blokované hodnoty (předvyplněné uživatelem)
BLOCK_VALUES = {"R", "DOV", "AMB", "GEN", "K", "COS", "C", "S", "POŽ"}
HOURS_FIXED = {"R": 8.0, "AMB": 8.0, "S": 8.0, "COS": 7.5, "K": 6.0}

MAX_NURSE_ROW = 46


def norm_text(s: str) -> str:
    """Normalizuje text - odstraní háčky, čárky, mezery"""
    s = (s or "").strip().upper()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return " ".join(s.split())


def to_float(v) -> float:
    """Převede hodnotu na float"""
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    return float(str(v).replace(",", "."))


def connect_to_sheets():
    """Připojí se k Google Sheets"""
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
    """Zjistí rok a měsíc z názvu listu"""
    mapping = {
        "LEDEN": 1, "UNOR": 2, "BREZEN": 3, "DUBEN": 4,
        "KVETEN": 5, "CERVEN": 6, "CERVENEC": 7, "SRPEN": 8,
        "ZARI": 9, "RIJEN": 10, "LISTOPAD": 11, "PROSINEC": 12,
    }
    
    name_norm = norm_text(sheet_name)
    for month_name, month_num in mapping.items():
        if month_name in name_norm:
            # Zkus najít rok v názvu (default 2026)
            year = 2026
            return year, month_num
    
    raise RuntimeError(f"Nerozumím názvu listu: '{sheet_name}'")


def detect_structure(ws_data):
    """Detekuje strukturu tabulky"""
    # Najdi hlavičku "Jméno"
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
    
    # Najdi plánovací sloupce (začíná číslem 1)
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
        raise RuntimeError("Nenašel jsem začátek měsíce (číslo 1)")
    
    return header_row, name_col, start_col


def load_employees(ws_data, header_row, name_col):
    """Načte seznam zaměstnanců"""
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
    """Načte fondy hodin z FONDY_HODIN"""
    try:
        ws = wb.worksheet("FONDY_HODIN")
        data = ws.get_all_values()
        
        # Hledej řádek pro tento měsíc
        for row in data[1:]:
            if len(row) < 3:
                continue
            
            # Zkus najít podle čísla měsíce v datu
            if "/" in row[0]:
                try:
                    parts = row[0].split("/")
                    date_month = int(parts[1]) if len(parts[0]) <= 2 else int(parts[0])
                    if date_month == month:
                        return to_float(row[1]), to_float(row[2])
                except:
                    pass
        
        print(f"⚠ Nenašel jsem fondy pro měsíc {month}, používám default")
        return 176.0, 165.0
    
    except Exception as e:
        print(f"⚠ Chyba při čtení FONDY_HODIN: {e}")
        return 176.0, 165.0


def load_employee_types(wb, employees):
    """Načte typy úvazků zaměstnanců"""
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
    except Exception as e:
        print(f"⚠ Chyba při čtení ZAMESTNANCI: {e}")
        return {}


def plan_shifts_v2(sheet_name: str):
    """
    Hlavní funkce - naplánuje směny pro daný list
    """
    
    print("=" * 60)
    print("Plánovač služeb V2")
    print("=" * 60)
    
    # Připoj se
    print(f"\n[1/7] Připojuji se k tabulce...")
    wb = connect_to_sheets()
    print(f"✓ Připojeno: {wb.title}")
    
    # Zjisti rok a měsíc z názvu listu
    print(f"\n[2/7] Zpracovávám list '{sheet_name}'...")
    year, month = get_month_from_sheet_name(sheet_name)
    days_in_month = calendar.monthrange(year, month)[1]
    print(f"✓ Rok: {year}, Měsíc: {month}, Dní: {days_in_month}")
    
    # Načti plánovací list
    print(f"\n[3/7] Načítám data...")
    ws = wb.worksheet(sheet_name)
    ws_data = ws.get_all_values()
    
    # Detekuj strukturu
    header_row, name_col, start_col = detect_structure(ws_data)
    plan_cols = list(range(start_col, start_col + days_in_month))
    print(f"✓ Struktura: hlavička={header_row}, jméno={name_col}, dny={start_col}–{start_col + days_in_month - 1}")
    
    # Načti zaměstnance
    print(f"\n[4/7] Načítám zaměstnance...")
    employees = load_employees(ws_data, header_row, name_col)
    print(f"✓ Nalezeno {len(employees)} zaměstnanců")
    
    if len(employees) == 0:
        raise RuntimeError("Žádní zaměstnanci!")
    
    # Najdi staniční sestru
    station_idx = None
    for i, emp in enumerate(employees):
        if "STARA" in norm_text(emp['name']):
            station_idx = i
            print(f"✓ Staniční sestra: {emp['name']} (řádek {emp['row']})")
            break
    
    if station_idx is None:
        print("⚠ Nenašel jsem 'Stará' - beru první osobu jako staniční")
        station_idx = 0
    
    # Načti fondy hodin
    print(f"\n[5/7] Načítám fondy hodin...")
    fond_1s, fond_05s = load_hours_fund(wb, year, month)
    print(f"✓ Fond: 1S={fond_1s}h, 0.5S={fond_05s}h")
    
    # Načti typy úvazků
    emp_types = load_employee_types(wb, employees)
    
    # Spočítej cílové hodiny
    for emp in employees:
        name_norm = norm_text(emp['name'])
        typ = emp_types.get(name_norm, "1S")
        if "1S" in typ:
            emp['target_hours'] = fond_1s * emp['uvazek']
        else:
            emp['target_hours'] = fond_05s * emp['uvazek']
    
    # Načti předvyplněné směny
    print(f"\n[6/7] Načítám předvyplněné směny...")
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
    
    # Doplň R pro staniční sestru (po–pá)
    for di in range(days_in_month):
        dt = datetime.date(year, month, di + 1)
        if dt.weekday() < 5:  # po–pá
            if fixed[station_idx][di] is None:
                fixed[station_idx][di] = "R"
                fixed_hours[station_idx] += HOURS_FIXED["R"]
    
    print(f"✓ Načteny předvyplněné směny")
    print(f"✓ Staniční má předvyplněno {fixed_hours[station_idx]}h")
    
    # SPUSŤ PLÁNOVÁNÍ
    print(f"\n[7/7] Spouštím plánování...")
    result = run_planner(employees, fixed, fixed_hours, days_in_month, year, month, station_idx)
    
    if not result:
        raise RuntimeError("Nelze najít řešení!")
    
    assign, hours = result
    
    # ZÁPIS DO TABULKY
    print(f"\n{'=' * 60}")
    print("Zapisuji výsledky...")
    print('=' * 60)
    
    # Připrav všechny změny do jednoho batch update
    write_count = 0
    
    # Najdi rozsah který potřebujeme
    min_row = min(e['row'] for e in employees)
    max_row = max(e['row'] for e in employees)
    min_col = min(plan_cols)
    max_col = max(plan_cols)
    
    # Vytvoř prázdnou mřížku
    num_rows = max_row - min_row + 1
    num_cols = max_col - min_col + 1
    grid = [[None] * num_cols for _ in range(num_rows)]
    
    # Naplň mřížku hodnotami
    for di, col_num in enumerate(plan_cols):
        for i, emp in enumerate(employees):
            row_data = ws_data[emp['row'] - 1] if (emp['row'] - 1) < len(ws_data) else []
            col_idx = col_num - 1
            
            # Vezmi původní hodnotu
            orig = ""
            if col_idx < len(row_data):
                orig = row_data[col_idx]
            
            # Zapiš jen pokud je prázdné NEBO je to staniční s R
            new_val = assign[i][di]
            if new_val and (orig in (None, "", 0) or (i == station_idx and new_val == "R")):
                grid_row = emp['row'] - min_row
                grid_col = col_num - min_col
                grid[grid_row][grid_col] = new_val
                write_count += 1
    
    # Zapiš celou mřížku najednou (1 API call!)
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
    
    return {
        "status": "success",
        "sheet": sheet_name,
        "employees": len(employees),
        "days": days_in_month,
        "written": write_count
    }


def run_planner(employees, fixed, fixed_hours, days, year, month, station_idx):
    """
    Backtracking plánovač s prioritou na vyrovnané hodiny
    """
    
    P = len(employees)
    D = days
    
    # Připrav struktury
    assign = [row[:] for row in fixed]
    hours = fixed_hours[:]
    target = [e['target_hours'] for e in employees]
    
    # Zjisti víkendy
    weekends = []
    for di in range(D):
        dt = datetime.date(year, month, di + 1)
        if dt.weekday() == 5:  # sobota
            weekends.append(di)
    
    def can_assign(i, di, shift):
        """Kontrola hard pravidel"""
        # Už tam něco je
        if assign[i][di] is not None:
            return False
        
        # Max 2 směny za sebou
        consec = 0
        for dj in range(max(0, di - MAX_CONSEC_SHIFTS), di):
            if assign[i][dj] in ("D", "N"):
                consec += 1
        if consec >= MAX_CONSEC_SHIFTS:
            return False
        
        # Po N musí volno
        if di > 0 and assign[i][di - 1] == "N":
            if shift in ("D", "N"):
                return False
        
        return True
    
    def score_person(i, shift):
        """
        Ohodnocení přiřazení směny osobě
        Čím MENŠÍ číslo, tím lepší
        """
        # Spočítej kolik bude mít po přiřazení
        future_hours = hours[i] + SHIFT_HOURS
        
        # HLAVNÍ KRITÉRIUM: rozdíl od targetu
        diff = abs(future_hours - target[i])
        
        # EXTRA penalizace za přesčas
        if future_hours > target[i]:
            diff *= 10.0  # Přesčas je 10x horší
        
        # Malý jitter pro randomizaci
        jitter = random.uniform(-0.01, 0.01)
        
        return diff + jitter
    
    def solve_day(di):
        """Rekurzivně naplánuj den"""
        if di == D:
            return True
        
        # Kolik D a N potřebujeme?
        needed_d = REQ_D
        needed_n = REQ_N
        
        for i in range(P):
            if assign[i][di] == "D":
                needed_d -= 1
            elif assign[i][di] == "N":
                needed_n -= 1
        
        # Už je hotovo
        if needed_d <= 0 and needed_n <= 0:
            return solve_day(di + 1)
        
        # Vytvoř seznam slotů
        slots = ["D"] * max(0, needed_d) + ["N"] * max(0, needed_n)
        
        def fill_slots(slot_idx, used):
            """Vyplň sloty"""
            if slot_idx == len(slots):
                return solve_day(di + 1)
            
            shift = slots[slot_idx]
            
            # Najdi kandidáty
            candidates = []
            for i in range(P):
                if i == station_idx:
                    continue
                if i in used:
                    continue
                if can_assign(i, di, shift):
                    score = score_person(i, shift)
                    candidates.append((score, i))
            
            # Seřaď podle skóre (nejlepší první)
            candidates.sort()
            
            # Zkus přiřadit
            for _, i in candidates:
                assign[i][di] = shift
                hours[i] += SHIFT_HOURS
                used.add(i)
                
                if fill_slots(slot_idx + 1, used):
                    return True
                
                # Undo
                assign[i][di] = None
                hours[i] -= SHIFT_HOURS
                used.remove(i)
            
            return False
        
        return fill_slots(0, set())
    
    # Najdi řešení
    print("Hledám řešení...")
    MAX_TRIES = 100
    
    for attempt in range(MAX_TRIES):
        # Reset
        assign = [row[:] for row in fixed]
        hours = fixed_hours[:]
        
        random.seed(year * 1000 + month * 100 + attempt)
        
        if solve_day(0):
            print(f"✓ Řešení nalezeno (pokus {attempt + 1})")
            return assign, hours
        
        if attempt % 10 == 0 and attempt > 0:
            print(f"  Pokus {attempt}...")
    
    return None


if __name__ == "__main__":
    # Test
    result = plan_shifts_v2("CERVEN")
    print("\n✓ HOTOVO!")
