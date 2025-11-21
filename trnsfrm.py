import xml.etree.ElementTree as ET
import pandas as pd
import os
import json

def xml_xls_to_json_filtered(xml_file_path, json_file_path):
    """
    Converte un foglio di calcolo XML Excel (con estensione .xls) in un file JSON.
    Vengono escluse le colonne indesiderate e rinominata la colonna 'data_inizio'.
    """
    try:
        # Colonne da ignorare, come richiesto
        COLUMNS_TO_IGNORE = [
            'Nota Agenda', 'data_fine', 'tutto_il_giorno', 
            'data_inserimento', 'classe_desc', 'gruppo_desc', 
            'aula', 'materia'
        ]
        
        # Rinominazione della colonna
        RENAME_MAP = {'data_inizio': 'data'}
        
        # --- 1. Parsing XML e Estrazione Intestazioni ---
        NAMESPACES = {
            'ss': 'urn:schemas-microsoft-com:office:spreadsheet',
            '': 'urn:schemas-microsoft-com:office:spreadsheet'
        }
        
        tree = ET.parse(xml_file_path)
        root = tree.getroot()
        data = []
        worksheet = root.find('.//Worksheet', NAMESPACES)

        # Estrae le intestazioni originali
        header_row = worksheet.find('.//Row', NAMESPACES)
        original_columns = [cell.find('ss:Data', NAMESPACES).text 
                            for cell in header_row.findall('ss:Cell', NAMESPACES) 
                            if cell.find('ss:Data', NAMESPACES) is not None]

        # Mappa per tenere traccia dell'indice originale delle colonne da includere
        column_indices_to_include = []
        final_columns = []
        
        for i, col_name in enumerate(original_columns):
            if col_name not in COLUMNS_TO_IGNORE:
                column_indices_to_include.append(i)
                
                # Applica la rinominazione se la colonna è 'data_inizio'
                if col_name in RENAME_MAP:
                    final_columns.append(RENAME_MAP[col_name])
                else:
                    final_columns.append(col_name)

        # --- 2. Estrazione Dati e Filtraggio ---
        # Itera su tutte le righe di dati, saltando la prima (l'intestazione)
        for row in worksheet.findall('ss:Table/ss:Row', NAMESPACES)[1:]:
            row_data_full = [cell.find('ss:Data', NAMESPACES).text 
                             for cell in row.findall('ss:Cell', NAMESPACES)]
            
            # Pulisce i dati della riga e li estende con None se troppo corti
            while len(row_data_full) < len(original_columns):
                row_data_full.append(None)
                
            # Filtra i dati della riga in base agli indici da includere
            filtered_row_data = [row_data_full[i] for i in column_indices_to_include]
            
            # Aggiunge i dati filtrati all'elenco principale
            data.append(filtered_row_data)

        # --- 3. Conversione a JSON ---
        # Crea un DataFrame solo con le colonne e i dati desiderati
        df = pd.DataFrame(data, columns=final_columns)

        # Converti il DataFrame in una stringa JSON
        json_data = df.to_json(orient="records", indent=4, date_format='iso', force_ascii=False)
        
        # Scrivi la stringa JSON nel file di output
        with open(json_file_path, 'w', encoding='utf-8') as f:
            f.write(json_data)
        
        print(f"✅ Conversione riuscita di '{os.path.basename(xml_file_path)}' in '{os.path.basename(json_file_path)}' con colonne filtrate e rinominate.")
        
    except FileNotFoundError:
        print(f"❌ Errore: Il file '{xml_file_path}' non è stato trovato.")
    except Exception as e:
        print(f"❌ Si è verificato un errore durante la conversione: {e}")


# --- Esempio di Utilizzo ---
# ASSICURATI DI CAMBIARE IL NOME DEL FILE CON QUELLO REALE
EXCEL_INPUT = 'downloads/data.xls' 
JSON_OUTPUT = 'output.json'

# Esegui la trasformazione
xml_xls_to_json_filtered(EXCEL_INPUT, JSON_OUTPUT)