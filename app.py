import streamlit as st
import openpyxl
from datetime import datetime
from collections import defaultdict
import pandas as pd
import io
import pytz

# --- Funzioni di Elaborazione Dati ---

def get_current_time_rome():
    rome_tz = pytz.timezone('Europe/Rome')
    return datetime.now(rome_tz)

def generate_forecasting_data(input_excel_file):
    # ... (codice di generate_forecasting_data invariato) ...
    try:
        workbook = openpyxl.load_workbook(input_excel_file)
        sheet = workbook["Sheet1"]
    except Exception as e:
        st.error(f"Errore durante l'apertura del file Excel: {e}")
        return None
    suppliers_data = defaultdict(lambda: {
        "name": "", "monthly_totals": defaultdict(float),
        "antecedenti_2025_total": 0.0, "yearly_total": 0.0
    })
    current_supplier_code = None
    current_supplier_name = None
    for row in sheet.iter_rows():
        col_a_value = row[0].value if len(row) > 0 else None
        col_b_value = row[1].value if len(row) > 1 else None
        col_d_value = row[3].value if len(row) > 3 else None
        col_m_value = row[12].value if len(row) > 12 else None
        if col_a_value == "Cod. fornitore":
            current_supplier_code = col_b_value
            current_supplier_name = col_d_value
            if current_supplier_code: suppliers_data[current_supplier_code]["name"] = current_supplier_name
        elif current_supplier_code and col_a_value and isinstance(col_a_value, (str, int, float)) and str(col_a_value).strip() not in ["Cod. fornitore", "Subtotale"] and col_d_value and col_m_value is not None:
            try:
                delivery_date = None
                if isinstance(col_d_value, datetime): delivery_date = col_d_value
                elif isinstance(col_d_value, str):
                    try: delivery_date = datetime.strptime(col_d_value, "%Y-%m-%d %H:%M:%S")
                    except ValueError: 
                        try: delivery_date = datetime.strptime(col_d_value, "%Y-%m-%d")
                        except ValueError: 
                            try: delivery_date = datetime.strptime(col_d_value, "%d/%m/%Y")
                            except ValueError: pass
                if delivery_date and delivery_date <= datetime(2025, 12, 31):
                    amount = float(str(col_m_value).replace(",", "."))
                    if delivery_date.year == 2025: suppliers_data[current_supplier_code]['monthly_totals'][delivery_date.strftime("%m")] += amount
                    elif delivery_date.year < 2025: suppliers_data[current_supplier_code]['antecedenti_2025_total'] += amount
                    suppliers_data[current_supplier_code]['yearly_total'] += amount
            except (ValueError, TypeError): pass
    return suppliers_data

def add_contropartita_data(report_data, anagrafica_file, contropartita_file):
    """Arricchisce i dati del report con la colonna Contropartita."""
    try:
        wb_anagrafica = openpyxl.load_workbook(anagrafica_file, data_only=True)
        sheet_anagrafica = wb_anagrafica["Sheet1"]
        wb_contropartita = openpyxl.load_workbook(contropartita_file, data_only=True)
        sheet_contropartita = wb_contropartita["Foglio1"]

        contropartita_map = {str(row[0]): row[2] for row in sheet_contropartita.iter_rows(min_row=2, values_only=True) if row[0] is not None}
        
        anagrafica_map = {}
        rows_iter = sheet_anagrafica.iter_rows()
        for row in rows_iter:
            if isinstance(row[0].value, str) and row[0].value.strip().lower() == "codice":
                codice_fornitore = row[1].value
                if codice_fornitore:
                    try:
                        next_row = next(rows_iter)
                        ch_rifer_conto_code = next_row[10].value # Colonna K
                        if ch_rifer_conto_code: anagrafica_map[str(codice_fornitore)] = str(ch_rifer_conto_code)
                    except StopIteration: break
        
        # Aggiunge i dati di contropartita al dizionario del report
        for code, data in report_data.items():
            ch_rifer_code = anagrafica_map.get(str(code))
            data["Contropartita"] = contropartita_map.get(ch_rifer_code, "") if ch_rifer_code else ""
        
        return report_data
    except Exception as e:
        st.warning(f"Errore nell'elaborazione della contropartita: {e}. Verrà mostrato il report base.")
        return report_data

# --- Applicazione Streamlit ---
st.set_page_config(page_title="Report Previsioni di Costo Economico", layout="wide")
st.title("📊 Report Previsioni di Costo Economico")

st.subheader("vetronaviglio s.r.l.")

uploaded_file = st.file_uploader("1. Carica il file `ordfor06.xlsx`", type=["xlsx"])

# Sezione opzionale per Contropartita
with st.expander("2. Aggiungi Contropartita (Opzionale)"):
    uploaded_anagrafica = st.file_uploader("Carica `anagrafica.xlsx`", type=["xlsx"])
    uploaded_contropartita = st.file_uploader("Carica `contropartita.xlsx`", type=["xlsx"])

if uploaded_file:
    st.success("File `ordfor06.xlsx` caricato.")
    
    suppliers_data = generate_forecasting_data(uploaded_file)

    # Aggiungi dati contropartita se i file sono stati caricati
    if uploaded_anagrafica and uploaded_contropartita:
        st.info("File anagrafica e contropartita caricati. Aggiungo colonna al report.")
        suppliers_data = add_contropartita_data(suppliers_data, uploaded_anagrafica, uploaded_contropartita)

    if suppliers_data:
        # ... (resto del codice per la visualizzazione e il download) ...
        all_supplier_names_raw = sorted([data["name"] for data in suppliers_data.values()])
        all_supplier_names_for_multiselect = ["Tutti"] + all_supplier_names_raw
        selected_supplier_names = st.multiselect("Filtra Fornitori", options=all_supplier_names_for_multiselect, default=["Tutti"])

        if "Tutti" in selected_supplier_names: filtered_suppliers_data = suppliers_data
        elif selected_supplier_names: filtered_suppliers_data = {c: d for c, d in suppliers_data.items() if d["name"] in selected_supplier_names}
        else: filtered_suppliers_data = {}

        report_rows = []
        sorted_suppliers = sorted(filtered_suppliers_data.items(), key=lambda item: item[1]['name'])
        italian_month_names = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"]
        
        # Costruisci le righe del DataFrame
        for code, data in sorted_suppliers:
            row_data = {"Fornitore": data["name"], "Codice Fornitore": code}
            # Inserisci Contropartita se esiste
            if "Contropartita" in data: row_data["Contropartita"] = data["Contropartita"]
            row_data["Antecedenti 2025"] = data["antecedenti_2025_total"]
            for month_num in range(1, 13): row_data[italian_month_names[month_num - 1]] = data["monthly_totals"][f"{month_num:02d}"]
            row_data["Totale Anno"] = data["yearly_total"]
            report_rows.append(row_data)
        
        df = pd.DataFrame(report_rows)

        # Riordina le colonne per avere Contropartita in terza posizione
        if "Contropartita" in df.columns:
            cols = df.columns.tolist()
            cols.insert(2, cols.pop(cols.index('Contropartita')))
            df = df[cols]

        st.dataframe(df.style.format({col: "{:,.2f} €" for col in df.columns if col not in ["Fornitore", "Codice Fornitore", "Contropartita"]}), use_container_width=True)

        # --- Logica per il download del file Excel --- 
        output_excel_buffer = io.BytesIO()
        df.to_excel(output_excel_buffer, index=False, sheet_name='Report Previsioni')
        output_excel_buffer.seek(0)

        st.download_button(
            label="📥 Scarica Report Excel",
            data=output_excel_buffer,
            file_name="forecasting_completo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Nessun dato generato.")

st.markdown(f"---")
st.info(f"Ultimo aggiornamento: {get_current_time_rome().strftime('%d/%m/%Y %H:%M:%S')}")