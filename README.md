## Progetto: Previsioni di Costo Economico e Ricavi

### Descrizione
Questo progetto ha lo scopo di generare un report di previsione dei costi economici e dei ricavi basato sugli ordini di acquisto. Estrapola i dati da un file Excel contenente gli ordini fornitori, li aggrega per mese e, opzionalmente, li arricchisce con i dati di contropartita.

### Script Principale: `generate_report.py`

#### Scopo
Lo script `generate_report.py` è lo strumento unico per la generazione del report. Esegue due compiti principali:
1.  Legge i dati degli ordini da `ordfor06.xlsx` per creare un report di previsione aggregato per fornitore e mese.
2.  Se rileva la presenza del file `conto.xlsx`, arricchisce automaticamente il report aggiungendo la colonna "Contropartita" in terza posizione, basandosi sulla ragione sociale del fornitore e filtrando i conti che iniziano con "50.10", "50.20", "52.10", "54.10".

#### Funzionalità principali:
- Estrazione e aggregazione degli ordini per fornitore, mese e anno.
- Calcolo dei totali e formattazione in valuta (€).
- **Arricchimento automatico (opzionale)**: Aggiunta della colonna "Contropartita" tramite mappatura da `conto.xlsx` con filtro sui prefissi dei conti.
- Gestione dei fusi orari per un timestamp di aggiornamento sempre corretto (fuso orario di Roma).

#### Come utilizzare lo script

**Prerequisiti:**
Assicurati di avere Python e le librerie necessarie installate:
```bash
pip install openpyxl pytz
```

**File di Input:**
- **Obbligatorio:** `ordfor06.xlsx` - Il file contenente i dati degli ordini.
- **Opzionale:**
    - `conto.xlsx` - Per la mappatura Fornitore -> Conto (ragione sociale in colonna E, conto in colonna F, filtrato per prefissi specifici).

**Esecuzione:**
Per generare il report, esegui un singolo comando dal terminale:
```bash
python generate_report.py
```
Lo script produrrà il report di base e, se trova i file opzionali, lo arricchirà automaticamente.

**File di Output:**
Lo script genera il file `forecasting.xlsx`. Se l'arricchimento è avvenuto, la colonna "Contropartita" si troverà in terza posizione.

### Applicazione Web: `app.py`

#### Scopo
L'applicazione `app.py` fornisce un'interfaccia web interattiva (basata su Streamlit) per eseguire la stessa logica di `generate_report.py` direttamente dal browser.

#### Funzionalità:
- Caricamento del file `ordfor06.xlsx`.
- **Sezione opzionale** per caricare `conto.xlsx` e aggiungere la colonna "Contropartita" (ragione sociale in colonna E, conto in colonna F, filtrato per prefissi specifici).
- Filtro interattivo per i fornitori.
- Visualizzazione del report a schermo.
- Download del report finale in formato Excel.
- **Titolo aggiornato**: Il titolo dell'applicazione è stato modificato in "📊 Vetronaviglio s.r.l. - Report Previsioni Costo Economico".

#### Come utilizzare l'applicazione

**Prerequisiti:**
```bash
pip install streamlit pandas openpyxl pytz
```

**Esecuzione:**
```bash
streamlit run app.py
```

### Data Ultimo Aggiornamento: 09/10/2025 13:00:00
