import openpyxl
from datetime import datetime
from collections import defaultdict

def generate_forecasting_report(input_filepath, output_filepath, sheet_name="Sheet1"):
    """
    Genera un report di previsione da un file Excel.

    Args:
        input_filepath (str): Il percorso del file Excel di input (ordfor06.xlsx).
        output_filepath (str): Il percorso in cui verrà salvato il report Excel di output.
        sheet_name (str): Il nome del foglio da leggere dal file di input.
    """
    try:
        workbook = openpyxl.load_workbook(input_filepath)
        sheet = workbook[sheet_name]
    except FileNotFoundError:
        print(f"Errore: Il file '{input_filepath}' non è stato trovato.")
        return
    except KeyError:
        print(f"Errore: Il foglio '{sheet_name}' non è stato trovato nel file.")
        return
    except Exception as e:
        print(f"Errore durante l'apertura del file Excel: {e}")
        return

    suppliers_data = defaultdict(lambda: {
        "name": "",
        "monthly_totals": defaultdict(float),
        "antecedenti_2025_total": 0.0,       # New field for orders before 2025
        "yearly_total": 0.0                  # Total for all relevant orders
    })

    current_supplier_code = None
    current_supplier_name = None

    # Itera attraverso le righe direttamente usando sheet.iter_rows()
    for row_index, row in enumerate(sheet.iter_rows(), start=1):
        # Get values from relevant columns
        # openpyxl rows are 1-indexed, columns are 0-indexed in row object
        # Assicurati che la riga abbia abbastanza celle prima di accedervi
        col_a_value = row[0].value if len(row) > 0 else None
        col_b_value = row[1].value if len(row) > 1 else None
        col_c_value = row[2].value if len(row) > 2 else None
        col_d_value = row[3].value if len(row) > 3 else None
        col_m_value = row[12].value if len(row) > 12 else None # Column M is index 12 (for controvalore)

        # Identify Supplier Header
        if col_a_value == "Cod. fornitore":
            current_supplier_code = col_b_value
            current_supplier_name = col_d_value
            if current_supplier_code: # Assicurati che il codice non sia None
                suppliers_data[current_supplier_code]["name"] = current_supplier_name
                # print(f"DEBUG: New Supplier Block - Code: {current_supplier_code}, Name: {current_supplier_name}")
            else: # If supplier code is missing, reset current supplier context
                current_supplier_code = None
                current_supplier_name = None
        
        # Identify Order Line
        elif current_supplier_code and col_a_value and \
             isinstance(col_a_value, (str, int, float)) and \
             str(col_a_value).strip() not in ["Cod. fornitore", "Subtotale"] and \
             col_d_value and col_m_value is not None:
            
            try:
                # Parse delivery date
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
                                # Add format for DD/MM/YYYY
                                delivery_date = datetime.strptime(col_d_value, "%d/%m/%Y")
                            except ValueError:
                                # If all parsing fails, skip this order line
                                pass
                
                # Filter orders based on delivery date (up to 31/12/2025)
                if delivery_date and delivery_date <= datetime(2025, 12, 31):
                    amount_str = str(col_m_value).replace(',', '.')
                    amount = float(amount_str)

                    if delivery_date.year == 2025:
                        delivery_month = delivery_date.strftime("%m")
                        suppliers_data[current_supplier_code]['monthly_totals'][delivery_month] += amount
                    elif delivery_date.year < 2025:
                        suppliers_data[current_supplier_code]['antecedenti_2025_total'] += amount
                    
                    suppliers_data[current_supplier_code]['yearly_total'] += amount
                    # print(f"DEBUG: Order Processed - Supplier: {current_supplier_code}, Date: {delivery_date.strftime('%Y-%m-%d')}, Amount: {amount}")
                # else:
                    # print(f"DEBUG: Order Skipped (Date Filter) - Row: {row_index}, Date: {col_d_value}")
            except (ValueError, TypeError) as e:
                # print(f"DEBUG: Order Skipped (Parsing Error) - Row: {row_index}, Error: {e}, Date: {col_d_value}, Amount: {col_m_value}")
                pass # Skip rows with invalid date or amount
        # else:
            # print(f"DEBUG: Row Skipped (Not Order Line) - Row: {row_index}, Col A: {col_a_value}, Current Supplier: {current_supplier_code}")

    # Create a new workbook for the report
    report_workbook = openpyxl.Workbook()
    report_sheet = report_workbook.active
    report_sheet.title = "Report Previsioni"

    # Write headers
    headers = ["Fornitore", "Codice Fornitore", "Antecedenti 2025"] + \
              [datetime(2025, m, 1).strftime("%B") for m in range(1, 13)] + \
              ["Totale Anno"]
    report_sheet.append(headers)

    # Sort suppliers by name (Column A in the output report)
    sorted_suppliers = sorted(suppliers_data.items(), key=lambda item: item[1]['name'])

    # Write data for each supplier
    for code, data in sorted_suppliers:
        row_data = [data["name"], code, data["antecedenti_2025_total"]] # Add antecedenti total
        
        for month_num in range(1, 13):
            month_str = f"{month_num:02d}"
            row_data.append(data["monthly_totals"][month_str])
        
        row_data.append(data["yearly_total"])
        report_sheet.append(row_data)

    # Apply number format to amount columns
    currency_format = '#,##0.00 "€"' # More explicit format for Euro

    # Columns C (Antecedenti 2025) to P (Totale Anno) are 1-indexed
    for col_idx in range(3, 17): 
        for row_idx in range(2, report_sheet.max_row + 1): # Start from row 2 (after headers)
            cell = report_sheet.cell(row=row_idx, column=col_idx)
            if isinstance(cell.value, (int, float)): # Only format if it's a number
                cell.number_format = currency_format

    # Add update date
    report_sheet.append([]) # Empty row for spacing
    report_sheet.append([f"Aggiornato al: {datetime.now().strftime('%d/%m/%Y')}"])

    # Salva il report
    try:
        report_workbook.save(output_filepath)
        print(f"Report generato con successo in '{output_filepath}'")
    except Exception as e:
        print(f"Errore durante il salvataggio del report: {e}")

# Esempio di utilizzo (assumendo che lo script sia eseguito dalla root del progetto)
input_excel_path = "C:\\progetti_stefano\\automations\\previsioni_costo_economico\\ordfor06.xlsx"
output_excel_path = "C:\\progetti_stefano\\automations\\previsioni_costo_economico\\forecasting.xlsx"
generate_forecasting_report(input_excel_path, output_excel_path)