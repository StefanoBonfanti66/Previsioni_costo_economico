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
- Formattazione delle colonne degli importi come valuta (€).

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

### Applicazione Streamlit: `app.py`

#### Scopo
L'applicazione `app.py` fornisce un'interfaccia web interattiva per generare il report di previsione. Permette agli utenti di caricare il file `ordfor06.xlsx` e visualizzare il report direttamente nel browser.

#### Funzionalità aggiuntive:
- **Caricamento file**: Permette di caricare il file `ordfor06.xlsx` tramite un'interfaccia utente.
- **Filtro Fornitori**: Include un filtro multiselect per selezionare uno o più fornitori da visualizzare nel report. Di default, tutti i fornitori sono selezionati tramite l'opzione "Tutti".
- **Nomi dei mesi in italiano**: I nomi dei mesi nelle colonne del report sono visualizzati in italiano.
- **Download Report**: Permette di scaricare il report generato in formato Excel (`forecasting.xlsx`).
- **Timestamp di aggiornamento**: Mostra un timestamp nell'interfaccia per indicare l'ultimo aggiornamento dell'applicazione.

#### Come utilizzare l'applicazione Streamlit

**Prerequisiti:**
Oltre a `openpyxl`, è necessaria la libreria `streamlit` e `pandas`.
```bash
pip install streamlit pandas
```

**Esecuzione in locale (per test):**
1. Assicurati di avere `app.py` e `requirements.txt` nella directory del progetto.
2. Apri il terminale nella directory del progetto e esegui:
```bash
streamlit run app.py
```

**Deploy su Streamlit Cloud:**
L'applicazione è configurata per il deploy su Streamlit Cloud. Dopo aver caricato i file `app.py` e `requirements.txt` sul repository GitHub, è possibile deployare l'app seguendo le istruzioni di Streamlit Cloud (selezionando il branch `master` e `app.py` come main file).

### Data Ultimo Aggiornamento: 02/10/2025 12:45:00