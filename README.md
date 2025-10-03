## Progetto: Previsioni di Costo Economico e Ricavi

### Descrizione
Questo progetto ha lo scopo di generare un report di previsione dei costi economici e dei ricavi basato sugli ordini di acquisto. Estrapola i dati da un file Excel contenente gli ordini fornitori e li aggrega per fornitore, suddividendoli per mese e includendo un totale annuale.

### Script: `generate_report.py`

#### Scopo
Lo script `generate_report.py` legge i dati degli ordini da un file Excel di input, elabora le informazioni relative a fornitori, date di consegna e importi, e genera un report strutturato in un nuovo file Excel.

#### Funzionalità principali:
- Estrazione del codice e del nome del fornitore.
- Identificazione della data di consegna e dell'importo (controvalore) per ciascun ordine.
- Aggregazione degli importi per mese per gli ordini con data di consegna nel 2025.
- Raggruppamento degli importi per gli ordini con data di consegna antecedente al 2025 in una colonna dedicata.
- Calcolo del totale annuale per ciascun fornitore.
- Ordinamento dei fornitori in ordine alfabetico per nome nel report finale.
- Formattazione delle colonne degli importi come valuta (€), arrotondando al numero intero più vicino.

#### Come utilizzare lo script

**Prerequisiti:**
Assicurati di avere Python installato sul tuo sistema. Avrai anche bisogno della libreria `openpyxl` per la gestione dei file Excel. Se non l'hai già installata, puoi farlo tramite pip:
```bash
pip install openpyxl
```

**File di Input:**
Lo script si aspetta un file Excel chiamato `ordfor06.xlsx` nella stessa directory dello script. Il file deve contenere i dati degli ordini nel foglio `Sheet1` con la seguente struttura:
- **Colonna A:** Numero d'ordine (utilizzato per identificare le righe d'ordine).
- **Colonna D:** Data di consegna (formati supportati: `YYYY-MM-DD HH:MM:SS`, `YYYY-MM-DD`, `DD/MM/YYYY`).
- **Colonna M:** Controvalore dell'ordine (importo numerico, può usare la virgola come separatore decimale).
- Le righe che iniziano con "Cod. fornitore" nella Colonna A sono considerate intestazioni di blocco fornitore, con il codice fornitore in Colonna B e il nome del fornitore in Colonna D.

**Esecuzione:**
1. Salva lo script `generate_report.py` nella directory del tuo progetto.
2. Assicurati che il file `ordfor06.xlsx` sia presente nella stessa directory.
3. Apri il terminale o il prompt dei comandi, naviga alla directory del progetto e esegui lo script:
```bash
python generate_report.py
```

**File di Output:**
Lo script genererà un nuovo file Excel chiamato `forecasting.xlsx` nella stessa directory. Questo file conterrà il report con la seguente struttura:
- **Colonna A:** Nome del Fornitore (ordinato alfabeticamente).
- **Colonna B:** Codice Fornitore.
- **Colonna C:** Antecedenti 2025 (totale degli ordini con data di consegna precedente al 2025).
- **Colonne D-O:** Totali mensili per l'anno 2025 (Gennaio-Dicembre).
- **Colonna P:** Totale Anno (somma di tutti gli importi rilevanti).

### Script: app.py (Elaborazione Contropartita)

#### Scopo
Questo script arricchisce il report `forecasting.xlsx` (generato da `generate_report.py`) aggiungendo la colonna "Contropartita" per ogni fornitore.

#### Logica di funzionamento
1.  **Lettura File**: Lo script legge tre file Excel: `forecasting.xlsx`, `anagrafica.xlsx`, e `contropartita.xlsx`.
2.  **Mappatura Contropartita**: Crea una mappa dei codici di contropartita leggendo `contropartita.xlsx` (Codice in colonna A, Valore finale in colonna C).
3.  **Mappatura Fornitori**: Analizza il file `anagrafica.xlsx`, che ha una struttura a blocchi, per creare una mappa tra il "Codice Fornitore" e il rispettivo codice di contropartita. Un blocco fornitore è identificato dalla parola "Codice" in colonna A.
    - Il **Codice Fornitore** viene letto dalla colonna B della riga "Codice".
    - Il **Codice Contropartita** associato viene letto dalla colonna K della riga successiva.
4.  **Aggiornamento Report**: Lo script modifica il file `forecasting.xlsx`, inserendo la colonna "Contropartita" in terza posizione (colonna C) e popolandola con i valori trovati tramite le mappature.

#### Come utilizzare lo script

**File di Input:**
- `forecasting.xlsx`: Il report generato da `generate_report.py`.
- `anagrafica.xlsx`: File anagrafico dei fornitori contenente i codici fornitore e i codici di contropartita.
- `contropartita.xlsx`: File di mappatura che traduce i codici di contropartita nei valori finali.

**Esecuzione:**
Assicurati che tutti i file di input siano presenti e chiusi. Esegui lo script dal terminale:
```bash
python app.py
```

**File di Output:**
Lo script modifica direttamente il file `forecasting.xlsx`, aggiungendo la colonna "Contropartita" in terza posizione.

### Data Ultimo Aggiornamento: 03/10/2025 12:00:00