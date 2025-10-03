import openpyxl
from datetime import datetime
from collections import defaultdict
import os
import pytz

def get_contropartita_data(suppliers_data, anagrafica_path, contropartita_path):
    """Arricchisce i dati dei fornitori con le informazioni sulla contropartita."""
    try:
        wb_anagrafica = openpyxl.load_workbook(anagrafica_path, data_only=True)
        sheet_anagrafica = wb_anagrafica["Sheet1"]
        wb_contropartita = openpyxl.load_workbook(contropartita_path, data_only=True)
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
                        if ch_rifer_conto_code:
                            anagrafica_map[str(codice_fornitore)] = str(ch_rifer_conto_code)
                    except StopIteration:
                        break
        
        for code, data in suppliers_data.items():
            ch_rifer_code = anagrafica_map.get(str(code))
            data["Contropartita"] = contropartita_map.get(ch_rifer_code, "") if ch_rifer_code else ""
        
        return suppliers_data, True # Ritorna True per indicare che l'arricchimento è avvenuto
    except Exception:
        return suppliers_data, False # In caso di errore, ritorna i dati originali

def generate_forecasting_report(input_filepath, output_filepath, sheet_name="Sheet1"):
    """Genera un report di previsione e, se possibile, lo arricchisce con la contropartita."""
    try:
        workbook = openpyxl.load_workbook(input_filepath)
        sheet = workbook[sheet_name]
    except Exception as e:
        return f"Errore durante l'apertura del file di input: {e}"

    # ... (Logica di estrazione dati da ordfor06.xlsx - invariata) ...
    suppliers_data = defaultdict(lambda: {
        "name": "", "monthly_totals": defaultdict(float),
        "antecedenti_2025_total": 0.0, "yearly_total": 0.0
    })
    current_supplier_code = None
    for row in sheet.iter_rows():
        col_a_value = row[0].value if len(row) > 0 else None
        col_b_value = row[1].value if len(row) > 1 else None
        col_d_value = row[3].value if len(row) > 3 else None
        col_m_value = row[12].value if len(row) > 12 else None
        if col_a_value == "Cod. fornitore":
            current_supplier_code = col_b_value
            if current_supplier_code: suppliers_data[current_supplier_code]["name"] = row[3].value
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

    # --- Integrazione Logica Contropartita ---
    contropartita_added = False
    anagrafica_path = "anagrafica.xlsx"
    contropartita_path = "contropartita.xlsx"
    if os.path.exists(anagrafica_path) and os.path.exists(contropartita_path):
        suppliers_data, contropartita_added = get_contropartita_data(suppliers_data, anagrafica_path, contropartita_path)

    # --- Scrittura del file Excel di output ---
    report_workbook = openpyxl.Workbook()
    report_sheet = report_workbook.active
    report_sheet.title = "Report Previsioni"

    headers = ["Fornitore", "Codice Fornitore"]
    if contropartita_added:
        headers.append("Contropartita")
    headers.extend(["Antecedenti 2025"] + ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"] + ["Totale Anno"])
    report_sheet.append(headers)

    sorted_suppliers = sorted(suppliers_data.items(), key=lambda item: item[1]['name'])

    for code, data in sorted_suppliers:
        row_data = [data["name"], code]
        if contropartita_added:
            row_data.append(data.get("Contropartita", ""))
        row_data.extend([data["antecedenti_2025_total"]] + [data["monthly_totals"][f"{m:02d}"] for m in range(1, 13)] + [data["yearly_total"]])
        report_sheet.append(row_data)

    # ... (formattazione e aggiunta timestamp) ...
    currency_format = '#,##0 "€"'
    start_col = 4 if contropartita_added else 3
    for col_idx in range(start_col, len(headers) + 1):
        for row_idx in range(2, report_sheet.max_row + 1):
            cell = report_sheet.cell(row=row_idx, column=col_idx)
            if isinstance(cell.value, (int, float)): cell.number_format = currency_format

    rome_tz = pytz.timezone('Europe/Rome')
    now_rome = datetime.now(rome_tz)
    timestamp_str = now_rome.strftime('%d/%m/%Y %H:%M:%S')
    report_sheet.append([])
    report_sheet.append([f"Aggiornato al: {timestamp_str}"])

    try:
        report_workbook.save(output_filepath)
        message = f"Report generato con successo in '{output_filepath}'"
        if contropartita_added: message += " con colonna 'Contropartita'."
        return message
    except Exception as e:
        return f"Errore durante il salvataggio del report: {e}"

if __name__ == "__main__":
    result = generate_forecasting_report("ordfor06.xlsx", "forecasting.xlsx")
    print(result)