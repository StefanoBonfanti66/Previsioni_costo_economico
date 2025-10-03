import openpyxl

def process_files():
    """
    Aggiunge la colonna 'Contropartita' in terza posizione al file forecasting.xlsx,
    gestendo eventuali colonne duplicate prima dell'esecuzione.
    """
    try:
        # Carica i workbook
        wb_forecasting = openpyxl.load_workbook("forecasting.xlsx")
        sheet_forecasting = wb_forecasting["Report Previsioni"]

        # --- Logica di Pulizia Integrata ---
        header = [cell.value for cell in sheet_forecasting[1]]
        indices = [i for i, col_name in enumerate(header, 1) if col_name == "Contropartita"]
        # Rimuove le colonne partendo dalla fine per non alterare gli indici
        for index in sorted(indices, reverse=True):
            sheet_forecasting.delete_cols(index)
        
        # Salva lo stato pulito prima di procedere
        wb_forecasting.save("forecasting.xlsx")

        # --- Inizio Logica Principale ---
        # Ricarica i file per avere uno stato pulito
        wb_forecasting = openpyxl.load_workbook("forecasting.xlsx")
        sheet_forecasting = wb_forecasting["Report Previsioni"]
        wb_anagrafica = openpyxl.load_workbook("anagrafica.xlsx", data_only=True)
        wb_contropartita = openpyxl.load_workbook("contropartita.xlsx", data_only=True)
        sheet_anagrafica = wb_anagrafica["Sheet1"]
        sheet_contropartita = wb_contropartita["Foglio1"]

        # 1. Crea mappa da contropartita.xlsx
        contropartita_map = {}
        for row in sheet_contropartita.iter_rows(min_row=2, values_only=True):
            if row[0] is not None:
                contropartita_map[str(row[0])] = row[2]

        # 2. Crea mappa da anagrafica.xlsx
        anagrafica_map = {}
        rows_iter = sheet_anagrafica.iter_rows()
        for row in rows_iter:
            cell_value = row[0].value
            if isinstance(cell_value, str) and cell_value.strip().lower() == "codice":
                codice_fornitore = row[1].value
                if codice_fornitore:
                    try:
                        next_row = next(rows_iter)
                        ch_rifer_conto_code = next_row[10].value # Colonna K
                        if ch_rifer_conto_code:
                            anagrafica_map[str(codice_fornitore)] = str(ch_rifer_conto_code)
                    except StopIteration:
                        break

        # 3. Aggiorna forecasting.xlsx
        header = [cell.value for cell in sheet_forecasting[1]]
        codice_fornitore_col_index = header.index("Codice Fornitore")

        contropartita_values = []
        for row_idx in range(1, sheet_forecasting.max_row + 1):
            if row_idx == 1:
                contropartita_values.append("Contropartita")
                continue
            
            codice_fornitore = sheet_forecasting.cell(row=row_idx, column=codice_fornitore_col_index + 1).value
            val = ""
            if codice_fornitore:
                ch_rifer_code = anagrafica_map.get(str(codice_fornitore))
                if ch_rifer_code:
                    val = contropartita_map.get(ch_rifer_code, "")
            contropartita_values.append(val)

        sheet_forecasting.insert_cols(3)
        for row_idx, val in enumerate(contropartita_values, 1):
            sheet_forecasting.cell(row=row_idx, column=3, value=val)

        wb_forecasting.save("forecasting.xlsx")
        
        return "Elaborazione completata. Lo script è ora idempotente."

    except Exception as e:
        return f"Si è verificato un errore imprevisto: {e}"

if __name__ == "__main__":
    result = process_files()
    print(result)