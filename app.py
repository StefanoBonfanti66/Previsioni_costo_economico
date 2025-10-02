import streamlit as st
import openpyxl
from datetime import datetime
from collections import defaultdict
import pandas as pd
import io

# --- Report Generation Logic (Adapted from generate_report.py) ---
def generate_forecasting_data(input_excel_file, sheet_name="Sheet1"):
    """
    Generates forecasting data from an Excel file, returning a dictionary of supplier data.
    """
    try:
        # openpyxl.load_workbook can take a file-like object directly
        workbook = openpyxl.load_workbook(input_excel_file)
        sheet = workbook[sheet_name]
        # st.write("DEBUG: Workbook loaded successfully.") # Removed debug
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

    # Debug: Show first few rows of raw data - Removed debug
    # st.subheader("DEBUG: Raw Data Sample (First 10 rows)")
    # raw_data_sample = []
    # for r_idx, row in enumerate(sheet.iter_rows(), start=1):
    #     if r_idx > 10: break
    #     raw_data_sample.append([cell.value for cell in row])
    # st.write(pd.DataFrame(raw_data_sample))

    for row_index, row in enumerate(sheet.iter_rows(), start=1):
        col_a_value = row[0].value if len(row) > 0 else None
        col_b_value = row[1].value if len(row) > 1 else None
        col_c_value = row[2].value if len(row) > 2 else None
        col_d_value = row[3].value if len(row) > 3 else None
        col_m_value = row[12].value if len(row) > 12 else None

        if col_a_value == "Cod. fornitore":
            current_supplier_code = col_b_value
            current_supplier_name = col_d_value
            if current_supplier_code:
                suppliers_data[current_supplier_code]["name"] = current_supplier_name
                # st.write(f"DEBUG: New Supplier Block - Code: {current_supplier_code}, Name: {current_supplier_name}") # Removed debug
        
        elif current_supplier_code and col_a_value and \
             isinstance(col_a_value, (str, int, float)) and \
             str(col_a_value).strip() not in ["Cod. fornitore", "Subtotale"] and \
             col_d_value and col_m_value is not None:
            
            try:
                delivery_date = None
                if isinstance(col_d_value, datetime):
                    delivery_date = col_d_value
                elif isinstance(col_d_value, str):
                    try:
                        delivery_date = datetime.strptime(col_d_value, "%Y-%m-%d %H:%M:%S")
                    except ValueError:
                        try:
                            delivery_date = datetime.strptime(col_d_value, "%Y-%m-%d")
                        except ValueError:
                            try:
                                delivery_date = datetime.strptime(col_d_value, "%d/%m/%Y")
                            except ValueError:
                                pass
                
                if delivery_date and delivery_date <= datetime(2025, 12, 31):
                    amount_str = str(col_m_value).replace(',', '.')
                    amount = float(amount_str)

                    if delivery_date.year == 2025:
                        delivery_month = delivery_date.strftime("%m")
                        suppliers_data[current_supplier_code]['monthly_totals'][delivery_month] += amount
                    elif delivery_date.year < 2025:
                        suppliers_data[current_supplier_code]['antecedenti_2025_total'] += amount
                    
                    suppliers_data[current_supplier_code]['yearly_total'] += amount
                    # st.write(f"DEBUG: Order Processed - Supplier: {current_supplier_code}, Date: {delivery_date.strftime('%Y-%m-%d')}, Amount: {amount}") # Removed debug
                # else:
                    # st.write(f"DEBUG: Order Skipped (Date Filter) - Row: {row_index}, Date: {col_d_value}") # Removed debug
            except (ValueError, TypeError) as e:
                # st.write(f"DEBUG: Order Skipped (Parsing Error) - Row: {row_index}, Error: {e}, Date: {col_d_value}, Amount: {col_m_value}") # Removed debug
                pass
        # else:
            # if row_index < 15: # Limit debug output - Removed debug
                # st.write(f"DEBUG: Row Skipped (Not Order Line) - Row: {row_index}, Col A: {col_a_value}, Current Supplier: {current_supplier_code}") # Removed debug
    
    # st.write("DEBUG: Final suppliers_data (sample):", {k: suppliers_data[k] for k in list(suppliers_data)[:2]}) # Removed debug
    return suppliers_data

# --- Streamlit App ---
st.set_page_config(page_title="Report Previsioni di Costo Economico", layout="wide")

st.title("📊 Report Previsioni di Costo Economico")
st.markdown("Carica il tuo file `ordfor06.xlsx` per generare il report di previsione.")

uploaded_file = st.file_uploader("Scegli un file Excel (ordfor06.xlsx)", type=["xlsx"])

if uploaded_file:
    st.success("File caricato con successo!")
    
    # Generate data
    suppliers_data = generate_forecasting_data(uploaded_file, sheet_name="Sheet1")

    if suppliers_data:
        st.subheader("Report Generato")

        # Prepare data for DataFrame
        report_rows = []
        sorted_suppliers = sorted(suppliers_data.items(), key=lambda item: item[1]['name'])

        for code, data in sorted_suppliers:
            row_data = {
                "Fornitore": data["name"],
                "Codice Fornitore": code,
                "Antecedenti 2025": data["antecedenti_2025_total"]
            }
            for month_num in range(1, 13):
                month_str = f"{month_num:02d}"
                month_name = datetime(2025, month_num, 1).strftime("%B")
                row_data[month_name] = data["monthly_totals"][month_str]
            row_data["Totale Anno"] = data["yearly_total"]
            report_rows.append(row_data)
        
        df = pd.DataFrame(report_rows)
        
        # Display DataFrame
        st.dataframe(df.style.format(
            {col: '#,##0.00 €' for col in df.columns if col not in ["Fornitore", "Codice Fornitore"]}
        ), use_container_width=True)

        # Download button for Excel file
        output_excel_buffer = io.BytesIO()
        
        # Create a new workbook for the report
        report_workbook = openpyxl.Workbook()
        report_sheet = report_workbook.active
        report_sheet.title = "Report Previsioni"

        # Write headers
        headers = ["Fornitore", "Codice Fornitore", "Antecedenti 2025"] + \
                  [datetime(2025, m, 1).strftime("%B") for m in range(1, 13)] + \
                  ["Totale Anno"]
        report_sheet.append(headers)

        # Write data for each supplier
        for code, data in sorted_suppliers:
            row_data = [data["name"], code, data["antecedenti_2025_total"]]
            for month_num in range(1, 13):
                month_str = f"{month_num:02d}"
                row_data.append(data["monthly_totals"][month_str])
            row_data.append(data["yearly_total"])
            report_sheet.append(row_data)

        # Apply number format to amount columns
        currency_format = '#,##0.00 "€"'
        for col_idx in range(3, 17): # Columns C to P (1-indexed)
            for row_idx in range(2, report_sheet.max_row + 1): # Start from row 2 (after headers)
                cell = report_sheet.cell(row=row_idx, column=col_idx)
                if isinstance(cell.value, (int, float)):
                    cell.number_format = currency_format

        # Add update date
        report_sheet.append([])
        report_sheet.append([f"Aggiornato al: {datetime.now().strftime('%d/%m/%Y')}"])

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
st.info(f"Ultimo aggiornamento: {datetime.now().strftime('%d/%m/%Y')}")
