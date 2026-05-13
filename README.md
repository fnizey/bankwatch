# Norsk Bankanalyse Dashboard

Bloomberg-inspirert analyseverktøy for norske sparebanker – 164 banker, 2015–2025, 6 analytiske moduler.

**Live demo:** `https://[ditt-brukernavn].github.io/bankanalyse/`

---

## Innhold

- [Hva er dette?](#hva-er-dette)
- [Moduler](#moduler)
- [Rask start (lokal)](#rask-start-lokal)
- [GitHub Pages – steg for steg](#github-pages--steg-for-steg)
- [Oppdatere data](#oppdatere-data)
- [Prosjektstruktur](#prosjektstruktur)
- [Datakilder](#datakilder)
- [Teknisk stack](#teknisk-stack)

---

## Hva er dette?

Et profesjonelt, nettleserbasert dashbord for analyse av norske sparebanker. Ingen backend kreves – all analyse kjøres i nettleseren (JavaScript + Plotly.js) mot forhåndsbehandlede JSON-datafiler.

**Datasett:**
- 1 045 bank-år-observasjoner
- 164 unike banker
- 11 år (2015–2025)
- 40+ nøkkeltall per bank per år
- Norges Bank styringsrente 2000–2026 + fremskrivninger

---

## Moduler

| Modul | Beskrivelse |
|---|---|
| **Bankdashboard** | Historisk utvikling, KPI-kort, radar, benchmark mot sektor, CSV-eksport |
| **Sektoranalyse** | Gjennomsnitt, median, spredning, korrelasjonsmatrise, top-15 ranking |
| **Segmentanalyse** | Små / mellomstore / store banker, rentesensitivitet, volatilitet |
| **Regresjon** | OLS, rolling regression, scatter + residualplot, automatisk tolkning |
| **Monte Carlo** | Fan charts, 4 rentescenarioer, sannsynlighetsbarrer, 10 000 simuleringer |
| **Sammenligning** | Bank vs. bank, vs. sektor, vs. segment, Z-score, percentilplassering |

---

## Rask start (lokal)

### 1. Klon eller last ned prosjektet

```bash
git clone https://github.com/[ditt-brukernavn]/bankanalyse.git
cd bankanalyse
```

### 2. Installer Python-avhengigheter

```bash
pip install -r requirements.txt
```

> Krever Python 3.9+. Sjekk med `python3 --version`.

### 3. Legg inn Excel-kildefilene

Kopier alle 11 Excel-filer fra Finanstilsynet til `excel_data/`-mappen:

```
excel_data/
├── sparebankenes-arsregnskaper---annual-accounts-for-savings-banks-2015__3_.xlsx
├── Sparebankenes_årsregnskaper_2016__4_.xlsx
├── Sparebankenes_årsregnskaper_2017__3_.xlsx
├── Sparebankenes_årsregnskaper_2018__3_.xlsx
├── Sparebankenes_årsregnskaper_2019__3_.xlsx
├── Sparebankenes_årsregnskaper_2020__3_.xlsx
├── Sparebankenes_årsregnskaper_2021__3_.xlsx
├── Sparebankenes_årsregnskaper_2022__3_.xlsx
├── Sparebankenes_årsregnskaper_2023__7_.xlsx
├── Regnskapstall_for_norske_banker_for_året_2024__2_.xlsx
└── Regnskapstall_for_norske_banker_for_året_2025.xlsx
```

> **NB:** Excel-filene er ikke inkludert i repoet (for store for GitHub). Last dem ned fra Finanstilsynet og Norges Bank (se [Datakilder](#datakilder)).

### 4. Kjør dataprosessering

```bash
python3 scripts/process_data.py
```

Dette:
- Leser alle 11 Excel-filer
- Skraper Norges Banks rentebane direkte fra `norges-bank.no`
- Eksporterer `data/banks_data.json`, `data/policy_rates.json`, `data/summary.json`

### 5. Start lokal webserver

```bash
python3 -m http.server 8080
```

Åpne `http://localhost:8080` i nettleseren.

---

## GitHub Pages – steg for steg

### Forutsetninger
- GitHub-konto ([github.com](https://github.com))
- Git installert ([git-scm.com](https://git-scm.com))

---

### Steg 1: Opprett nytt GitHub-repo

1. Gå til [github.com/new](https://github.com/new)
2. Navn: `bankanalyse` (eller valgfritt)
3. Sett til **Public** (kreves for gratis GitHub Pages)
4. **Ikke** initialiser med README
5. Klikk **Create repository**

---

### Steg 2: Konfigurer `.gitignore`

Lag en fil `.gitignore` i prosjektmappen:

```
excel_data/
__pycache__/
*.pyc
.DS_Store
venv/
```

> Excel-filene holdes utenfor repoet – de er store og proprietære.

---

### Steg 3: Push prosjektet

```bash
cd bankanalyse

# Initialiser git (hvis ikke allerede gjort)
git init

# Legg til alle filer
git add .

# Første commit
git commit -m "Initial commit: Norwegian bank analysis dashboard"

# Koble til GitHub (bytt ut URL)
git remote add origin https://github.com/[ditt-brukernavn]/bankanalyse.git

# Push
git branch -M main
git push -u origin main
```

---

### Steg 4: Aktiver GitHub Pages

1. Gå til repoet på GitHub
2. Klikk **Settings** (tannhjulikonet)
3. Scroll ned til **Pages** i venstremenyen
4. Under **Source**: velg **Deploy from a branch**
5. Branch: **main**, mappe: **/ (root)**
6. Klikk **Save**

---

### Steg 5: Vent og åpne

GitHub Pages bruker 1–3 minutter på første deploy. Dashbordet er tilgjengelig på:

```
https://[ditt-brukernavn].github.io/bankanalyse/
```

Under **Settings → Pages** ser du eksakt URL og deploy-status.

---

### Verifiser at alt fungerer

Åpne nettleser-konsollen (F12 → Console). Du skal **ikke** se feil som:

```
Failed to fetch banks_data.json
```

Hvis du ser det, betyr det at `data/`-mappen mangler JSON-filene. Forsikre deg om at du har committed `data/*.json` (ikke bare `index.html`).

Sjekk at disse filene finnes i repoet ditt på GitHub:
```
data/banks_data.json    (~900 KB)
data/policy_rates.json  (~10 KB)
data/summary.json       (~1 KB)
index.html
```

---

## Oppdatere data

Når Finanstilsynet publiserer ny årsrapport (typisk april–mai):

1. Last ned ny Excel-fil til `excel_data/`
2. Kjør `python3 scripts/process_data.py`
3. Commit og push:

```bash
git add data/
git commit -m "Oppdater data: legger til [år]"
git push
```

GitHub Pages publiserer automatisk innen 1–2 minutter.

---

## Prosjektstruktur

```
bankanalyse/
├── index.html                  # Hele frontenden (self-contained)
├── requirements.txt            # Python-avhengigheter
├── README.md                   # Denne filen
├── .gitignore
│
├── data/                       # Genererte JSON-datafiler (committed)
│   ├── banks_data.json         # 1045 bank-år-poster, alle nøkkeltall
│   ├── policy_rates.json       # NB styringsrente 2000–2026 + prognose
│   └── summary.json            # Metadata: antall banker, år, etc.
│
├── scripts/
│   └── process_data.py         # Dataprosessering + NB-skraping
│
└── excel_data/                 # Excel-kildefiler (IKKE committed)
    └── [11 Excel-filer her]
```

---

## Datakilder

| Kilde | URL | Innhold |
|---|---|---|
| Finanstilsynet | [finanstilsynet.no](https://www.finanstilsynet.no/statistikk/bank-og-finansieringsforetak/sparebanker/) | Årsregnskaper 2015–2023 |
| Finanstilsynet | [finanstilsynet.no](https://www.finanstilsynet.no/statistikk/bank-og-finansieringsforetak/banker-og-kredittforetak/) | Regnskapstall 2024–2025 |
| Norges Bank | [norges-bank.no](https://www.norges-bank.no/en/topics/monetary-policy/policy-rate/) | Styringsrente historikk + prognose |

---

## Teknisk stack

| Komponent | Teknologi |
|---|---|
| Frontend | Vanilla JavaScript (ES6+) |
| Visualisering | [Plotly.js 2.27](https://plotly.com/javascript/) |
| Statistikk | [simple-statistics 7.8.3](https://simplestatistics.org/) |
| Styling | Custom CSS (Bloomberg dark theme) |
| Dataprosessering | Python 3 + pandas + openpyxl |
| NB-dataskraping | requests + BeautifulSoup4 |
| Hosting | GitHub Pages (statisk, ingen backend) |

**Ingen server kreves.** Alt kjøres i nettleseren mot statiske JSON-filer.

---

## Lisens

MIT – fritt til bruk, modifikasjon og distribusjon.

---

*Sist oppdatert: Mai 2026. Data: Finanstilsynet / Norges Bank.*
