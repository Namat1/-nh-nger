import streamlit as st
import pandas as pd
from io import BytesIO
import re

# Titel der App
st.title("Zulage GGL + Anhänger")

# Mehrere Dateien hochladen
uploaded_files = st.file_uploader("Lade deine Excel- oder CSV-Dateien hoch", type=["xlsx", "xls", "csv"], accept_multiple_files=True)

if uploaded_files:
    all_results = []  # Liste, um Ergebnisse zu speichern

    for uploaded_file in uploaded_files:
        try:
            # Dateiname extrahieren
            file_name = uploaded_file.name
            st.write(f"Verarbeite Datei: {file_name}")

            # Kalenderwoche aus dem Dateinamen extrahieren
            kw_match = re.search(r'KW(\d{1,2})', file_name, re.IGNORECASE)
            kalenderwoche = f"KW{kw_match.group(1)}" if kw_match else "Keine KW gefunden"

            # Datei lesen
            if uploaded_file.name.endswith(('.xlsx', '.xls')):
                df = pd.read_excel(uploaded_file, sheet_name="Touren")
                st.success(f"Das Blatt 'Touren' aus {file_name} wurde erfolgreich geladen!")
            else:
                df = pd.read_csv(uploaded_file)
                st.success(f"CSV-Datei {file_name} wurde erfolgreich geladen!")

            # Automatische Suchoptionen
            search_numbers = ["602", "620", "350", "520", "156"]
            search_strings = ["AZ", "Az", "az", "MW", "Mw", "mw"]

            # Benötigte Spalten prüfen
            required_columns = ['Unnamed: 0', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 6',
                                'Unnamed: 7', 'Unnamed: 11', 'Unnamed: 12', 'Unnamed: 14']

            if all(col in df.columns for col in required_columns):
                # Spalten als Strings behandeln
                df['Unnamed: 11'] = df['Unnamed: 11'].astype(str)
                df['Unnamed: 14'] = df['Unnamed: 14'].astype(str)

                # Filter nach Suchoptionen
                number_matches = df[
                    df['Unnamed: 11'].isin(search_numbers) & 
                    (df['Unnamed: 11'] != "607")
                ]
                text_matches = df[
                    df['Unnamed: 14'].str.contains('|'.join(search_strings), case=False, na=False) &
                    (df['Unnamed: 11'] != "607")
                ]
                combined_results = pd.concat([number_matches, text_matches]).drop_duplicates()

                # Spalten extrahieren und umbenennen
                renamed_columns = {
                    'Unnamed: 0': 'Tour',
                    'Unnamed: 3': 'Nachname',
                    'Unnamed: 4': 'Vorname',
                    'Unnamed: 6': 'Nachname 2',
                    'Unnamed: 7': 'Vorname 2',
                    'Unnamed: 11': 'Kennzeichen',
                    'Unnamed: 12': 'Gz / GGL',
                    'Unnamed: 14': 'Art 2'
                }
                final_results = combined_results[required_columns].rename(columns=renamed_columns)

                # Sortieren und Verdienst berechnen
                final_results = final_results.sort_values(by=['Nachname', 'Vorname'])
                payment_mapping = {"602": 40, "156": 40, "620": 20, "350": 20, "520": 20}

                def calculate_payment(row):
                    kennzeichen = row['Kennzeichen']
                    art_2 = row['Art 2'].strip().upper()
                    return payment_mapping.get(kennzeichen, 0) if art_2 == "AZ" else 0

                # Verdienst berechnen
                final_results['Verdienst'] = final_results.apply(calculate_payment, axis=1)

                # Zeilen mit 0 oder NaN in "Verdienst" entfernen
                final_results = final_results[(final_results['Verdienst'] > 0) & final_results['Verdienst'].notna()]

                # KW zur Ergebnis-Tabelle hinzufügen
                final_results['KW'] = kalenderwoche

                # Ergebnisse sammeln
                all_results.append(final_results)

        except Exception as e:
            st.error(f"Fehler beim Verarbeiten der Datei {file_name}: {e}")

    # Gesamtergebnisse zusammenführen
    if all_results:
        combined_results = pd.concat(all_results, ignore_index=True)

        # Ergebnisse vorbereiten
        combined_results['Verdienst'] = (
            combined_results['Verdienst']
            .replace(r'[^\d.]', '', regex=True)
            .astype(float)
        )
        combined_results.fillna(0, inplace=True)

        # Ergebnisse in eine Excel-Datei exportieren
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            combined_results.to_excel(writer, index=False, sheet_name="Suchergebnisse", startrow=1)

            workbook = writer.book
            worksheet = writer.sheets["Suchergebnisse"]

            # Kopfzeilenformat
            header_format = workbook.add_format({'bold': True, 'bg_color': '#D9EAD3', 'border': 1})
            blue_format = workbook.add_format({'bg_color': '#DDEBF7', 'border': 1})
            white_format = workbook.add_format({'bg_color': '#FFFFFF', 'border': 1})

            # Kopfzeile formatieren
            for col_num, value in enumerate(combined_results.columns):
                worksheet.write(0, col_num, value, header_format)

            # Zeilen formatieren
            for row_num, row_data in enumerate(combined_results.values, start=1):
                row_format = blue_format if row_num % 2 == 0 else white_format
                for col_num, cell_data in enumerate(row_data):
                    if isinstance(cell_data, (int, float)):
                        worksheet.write_number(row_num, col_num, cell_data, row_format)
                    elif isinstance(cell_data, str):
                        worksheet.write_string(row_num, col_num, cell_data, row_format)
                    elif cell_data is None:
                        worksheet.write_blank(row_num, col_num, None, row_format)
                    else:
                        worksheet.write_string(row_num, col_num, str(cell_data), row_format)

            # Auto-Spaltenbreite anpassen
            for col_num, column_name in enumerate(combined_results.columns):
                max_width = max(
                    len(str(column_name)),
                    combined_results[column_name].astype(str).map(len).max()
                )
                worksheet.set_column(col_num, col_num, max_width + 2)

            output.seek(0)

        # Download-Button anzeigen
        st.download_button(
            label="Kombinierte Ergebnisse als Excel herunterladen",
            data=output.getvalue(),
            file_name="Suchergebnisse_format.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
