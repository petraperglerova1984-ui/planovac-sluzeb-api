# Plánovač služeb - API server

Python server pro automatické plánování služeb v Google Sheets.

## Struktura projektu

```
planovac_projekt/
├── app.py                  # Flask API server
├── planner_sheets.py       # Hlavní plánovací logika
├── credentials.json        # Google Service Account credentials
├── requirements.txt        # Python závislosti
├── Procfile               # Pro deployment na Render
└── README.md              # Tento soubor
```

## Lokální spuštění

1. Nainstaluj závislosti:
```bash
pip install -r requirements.txt
```

2. Spusť server:
```bash
python app.py
```

Server poběží na `http://localhost:5000`

## API Endpointy

### GET /
Vrátí status API

### GET /health
Health check endpoint

### POST /plan
Spustí plánování služeb

Příklad requestu:
```bash
curl -X POST http://localhost:5000/plan
```

## Deployment na Render.com

1. Vytvoř nový Web Service na render.com
2. Připoj GitHub repository
3. Nastav:
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `gunicorn app:app`
4. Přidej environment proměnné (pokud potřeba)

## Google Sheets API

Server používá Service Account pro přístup k Google Sheets.
Ujisti se, že:
1. Máš vytvořený Service Account v Google Cloud Console
2. Stažený JSON klíč je uložený jako `credentials.json`
3. Google Sheets tabulka je sdílená s emailem Service Accountu (s právy Editor)

## Konfigurace

V souboru `planner_sheets.py` na začátku:
- `SPREADSHEET_ID` - ID tvé Google Sheets tabulky
- `CREDENTIALS_FILE` - cesta k JSON klíči

## Plánovací pravidla

**Hard pravidla:**
- 3× denní směna (D) a 3× noční (N) každý den
- Max 3 směny za sebou
- Po noční nesmí hned další směna
- Staniční sestra má po–pá "R"

**Soft pravidla:**
- Férové rozložení směn
- Min. 2 volné víkendy
- Optimalizace vzoru DN00

## Podpora

Pro problémy nebo dotazy kontaktuj autora.
