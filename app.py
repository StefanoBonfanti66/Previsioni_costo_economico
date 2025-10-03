import streamlit as st
import openpyxl
from datetime import datetime
from collections import defaultdict
import pandas as pd
import io
import pytz # Aggiunto per la gestione dei fusi orari

# --- Funzione per ottenere l'orario corretto ---
def get_current_time_rome():
    """Restituisce l'orario attuale nel fuso orario di Roma."""
    rome_tz = pytz.timezone('Europe/Rome')
    return datetime.now(rome_tz)

# --- Logica di Generazione Report (Adattata da generate_report.py) ---
def generate_forecasting_data(input_excel_file, sheet_name="Sheet1"):
    """
    Genera i dati di previsione da un file Excel.
    """
    try:
        workbook = openpyxl.load_workbook(input_excel_file)
        sheet = workbook[sheet_name]
    except Exception as e:
        st.error(f"Errore durante l'apertura del file Excel: {e}")
        return None

    suppliers_data = defaultdict(lambda: {
        "name": "",
        "monthly_totals": defaultdict(float),
        "antecedenti_2025_total": 0.0,
        "yearly_total": 0.0
    })

    current_supplier_code = None
    current_supplier_name = None

    for row_index, row in enumerate(sheet.iter_rows(), start=1):
        col_a_value = row[0].value if len(row) > 0 else None
        col_b_value = row[1].value if len(row) > 1 else None
        col_d_value = row[3].value if len(row) > 3 else None
        col_m_value = row[12].value if len(row) > 12 else None

        if col_a_value == "Cod. fornitore":
            current_supplier_code = col_b_value
            current_supplier_name = col_d_value
            if current_supplier_code:
                suppliers_data[current_supplier_code]["name"] = current_supplier_name
        
        elif current_supplier_code and col_a_value and \
             isinstance(col_a_value, (str, int, float)) and \
             str(col_a_value).strip() not in ["Cod. fornitore", "Subtotale"] and \
             col_d_value and col_m_value is not None:
            
            try:
                delivery_date = None
                if isinstance(col_d_value, datetime):
                    delivery_date = col_d_value
                elif isinstance(col_d_value, str):
                    try: delivery_date = datetime.strptime(col_d_value, "%Y-%m-%d %H:%M:%S")
                    except ValueError: 
                        try: delivery_date = datetime.strptime(col_d_value, "%Y-%m-%d")
                        except ValueError: 
                            try: delivery_date = datetime.strptime(col_d_value, "%d/%m/%Y")
                            except ValueError: pass
                
                if delivery_date and delivery_date <= datetime(2025, 12, 31):
                    amount_str = str(col_m_value).replace(",", ".")
                    amount = float(amount_str)

                    if delivery_date.year == 2025:
                        suppliers_data[current_supplier_code]['monthly_totals'][delivery_date.strftime("%m")] += amount
                    elif delivery_date.year < 2025:
                        suppliers_data[current_supplier_code]['antecedenti_2025_total'] += amount
                    
                    suppliers_data[current_supplier_code]['yearly_total'] += amount
            except (ValueError, TypeError): pass
    return suppliers_data

# --- Applicazione Streamlit ---
st.set_page_config(page_title="Report Previsioni di Costo Economico", layout="wide")

st.title("📊 Report Previsioni di Costo Economico")
st.markdown("Carica il tuo file `ordfor06.xlsx` per generare il report di previsione.")

uploaded_file = st.file_uploader("Scegli un file Excel (ordfor06.xlsx)", type=["xlsx"])

if uploaded_file:
    st.success("File caricato con successo!")
    
    suppliers_data = generate_forecasting_data(uploaded_file, sheet_name="Sheet1")

    if suppliers_data:
        st.subheader("Report Generato")

        all_supplier_names_raw = sorted([data["name"] for data in suppliers_data.values()])
        all_supplier_names_for_multiselect = ["Tutti"] + all_supplier_names_raw
        
        selected_supplier_names = st.multiselect(
            "Seleziona Fornitori",
            options=all_supplier_names_for_multiselect,
            default=["Tutti"]
        )

        if "Tutti" in selected_supplier_names:
            filtered_suppliers_data = suppliers_data 
        elif selected_supplier_names:
            filtered_suppliers_data = {code: data for code, data in suppliers_data.items() if data["name"] in selected_supplier_names}
        else:
            filtered_suppliers_data = {}

        report_rows = []
        sorted_suppliers = sorted(filtered_suppliers_data.items(), key=lambda item: item[1]['name'])

        italian_month_names = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"]

        for code, data in sorted_suppliers:
            row_data = {"Fornitore": data["name"], "Codice Fornitore": code, "Antecedenti 2025": data["antecedenti_2025_total"]}
            for month_num in range(1, 13):
                row_data[italian_month_names[month_num - 1]] = data["monthly_totals"].get(f"{month_num:02d}", 0.0)
            row_data["Totale Anno"] = data["yearly_total"]
            report_rows.append(row_data)
        
        df = pd.DataFrame(report_rows)
        
        st.dataframe(df.style.format({col: "{:,.2f} €" for col in df.columns if col not in ["Fornitore", "Codice Fornitore"]}), use_container_width=True)

        output_excel_buffer = io.BytesIO()
        report_workbook = openpyxl.Workbook()
        report_sheet = report_workbook.active
        report_sheet.title = "Report Previsioni"

        headers = ["Fornitore", "Codice Fornitore", "Antecedenti 2025"] + italian_month_names + ["Totale Anno"]
        report_sheet.append(headers)

        for code, data in sorted_suppliers:
            row_data = [data["name"], code, data["antecedenti_2025_total"]]
            for month_num in range(1, 13):
                row_data.append(data["monthly_totals"].get(f"{month_num:02d}", 0.0))
            row_data.append(data["yearly_total"])
            report_sheet.append(row_data)

        currency_format = '#,##0.00 "€"'
        for col_idx in range(3, 17):
            for row_idx in range(2, report_sheet.max_row + 1):
                cell = report_sheet.cell(row=row_idx, column=col_idx)
                if isinstance(cell.value, (int, float)):
                    cell.number_format = currency_format

        # --- CORREZIONE FUSO ORARIO ---
        now_rome_str = get_current_time_rome().strftime('%d/%m/%Y %H:%M:%S')
        report_sheet.append([])
        report_sheet.append([f"Aggiornato al: {now_rome_str}"])

        report_workbook.save(output_excel_buffer)
        output_excel_buffer.seek(0)

        st.download_button(
            label="Scarica Report Excel",
            data=output_excel_buffer,
            file_name="forecasting.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Nessun dato generato. Controlla il formato del file o i dati.")

st.markdown(f"---")
# --- CORREZIONE FUSO ORARIO ---
st.info(f"Ultimo aggiornamento: {get_current_time_rome().strftime('%d/%m/%Y %H:%M:%S')}")
