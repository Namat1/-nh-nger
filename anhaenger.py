import streamlit as st
import pandas as pd
from io import BytesIO
import re

# Titel der App
st.title("Zulage GGL + Anhänger")

# Mehrere Dateien hochladen
uploaded_files = st.file_uploader("Lade deine Excel- oder CSV-Dateien hoch", type=["xlsx", "xls", "csv"], accept_multiple_files=True)

# Variablen zur Zwischenspeicherung
combined_results = None
combined_summary = None

if uploaded_files:
    all_results = []  # Liste, um Ergebnisse zu speichern
    all_summaries = []  # Liste, um Zusammenfassungen zu speichern
    
    progress_bar = st.progress(0)  # Fortschrittsbalken hinzufügen
    total_files = len(uploaded_files)

    for idx, uploaded_file in enumerate(uploaded_files):
        try:
            # Fortschrittsanzeige aktualisieren
            progress_bar.progress((idx + 1) / total_files)

            # Dateiname extrahieren
            file_name = uploaded_file.name

            # Kalenderwoche aus dem Dateinamen extrahieren
            kw_match = re.search(r'KW(\d{1,2})', file_name, re.IGNORECASE)
            kalenderwoche = f"KW{kw_match.group(1)}" if kw_match else "Keine KW gefunden"

            # Datei lesen
            if uploaded_file.name.endswith(('.xlsx', '.xls')):
                df = pd.read_excel(uploaded_file, sheet_name="Touren")
            else:
                df = pd.read_csv(uploaded_file)

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
                combined_results_df = pd.concat([number_matches, text_matches]).drop_duplicates()

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
                final_results = combined_results_df[required_columns].rename(columns=renamed_columns)

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

                # Euro-Zeichen in den Suchergebnissen hinzufügen
                final_results['Verdienst'] = final_results['Verdienst'].apply(lambda x: f"{x} €")

                # KW zur Ergebnis-Tabelle hinzufügen
                final_results['KW'] = kalenderwoche

                # Ergebnisse sammeln
                all_results.append(final_results)

                # Zusammenfassung erstellen
                summary = final_results.copy()
                summary['Verdienst'] = summary['Verdienst'].str.replace(" €", "", regex=False).astype(float)
                summary = summary.groupby(['KW', 'Nachname', 'Vorname']).agg({'Verdienst': 'sum'}).reset_index()

                # Euro-Zeichen hinzufügen in der Zusammenfassung
                summary['Gesamtverdienst'] = summary['Verdienst'].apply(lambda x: f"{x} €")
                summary = summary.drop(columns=['Verdienst'])

                # Zusammenfassung in die Sammlung einfügen
                all_summaries.append(summary)
        except Exception as e:
            st.error(f"Fehler beim Verarbeiten der Datei {file_name}: {e}")

    # Gesamtergebnisse zusammenführen
    if all_results:
        combined_results = pd.concat(all_results, ignore_index=True)
        combined_summary = pd.concat(all_summaries, ignore_index=True)

        # Sortierung der Zusammenfassung nach numerischer KW
        combined_summary['KW_Numeric'] = combined_summary['KW'].str.extract(r'(\d+)').astype(int)
        combined_summary = combined_summary.sort_values(by=['KW_Numeric', 'Nachname', 'Vorname']).drop(columns=['KW_Numeric'])

    # Fortschrittsanzeige schließen und "FERTIG" anzeigen
    progress_bar.empty()  # Fortschrittsbalken entfernen
    st.success("FERTIG! Alle Dateien wurden verarbeitet.")

# Download-Bereich
if combined_results is not None and combined_summary is not None:
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Suchergebnisse
        search_sheet = writer.book.add_worksheet("Suchergebnisse")
        combined_results.to_excel(writer, sheet_name="Suchergebnisse", index=False, startrow=1)

        # Formatierungen hinzufügen
        header_format = writer.book.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
        blue_format = writer.book.add_format({'bg_color': '#E3F2FD', 'border': 1})
        green_format = writer.book.add_format({'bg_color': '#E8F5E9', 'border': 1})

        # Suchergebnisse - Kopfzeile formatieren
        for col_num, column_name in enumerate(combined_results.columns):
            search_sheet.write(0, col_num, column_name, header_format)

        # Suchergebnisse - Zeilen farbig formatieren
        current_format = blue_format
        for row_num in range(len(combined_results)):
            current_format = green_format if current_format == blue_format else blue_format
            for col_num in range(len(combined_results.columns)):
                search_sheet.write(row_num + 1, col_num, combined_results.iloc[row_num, col_num], current_format)

        # Automatische Spaltenbreite für Suchergebnisse
        for col_num, column_name in enumerate(combined_results.columns):
            max_content_width = max(
                combined_results[column_name].astype(str).apply(len).max(),
                len(column_name)
            )
            search_sheet.set_column(col_num, col_num, max_content_width + 2)

        # Zusammenfassung
        summary_sheet = writer.book.add_worksheet("Zusammenfassung")
        combined_summary.to_excel(writer, sheet_name="Zusammenfassung", index=False, startrow=1)

        # Zusammenfassung - Kopfzeile formatieren
        for col_num, column_name in enumerate(combined_summary.columns):
            summary_sheet.write(0, col_num, column_name, header_format)

        # Zusammenfassung - Zeilen farbig formatieren
        current_format = blue_format
        for row_num in range(len(combined_summary)):
            current_format = green_format if current_format == blue_format else blue_format
            for col_num in range(len(combined_summary.columns)):
                summary_sheet.write(row_num + 1, col_num, combined_summary.iloc[row_num, col_num], current_format)

        # Automatische Spaltenbreite für Zusammenfassung
        for col_num, column_name in enumerate(combined_summary.columns):
            max_content_width = max(
                combined_summary[column_name].astype(str).apply(len).max(),
                len(column_name)
            )
            summary_sheet.set_column(col_num, col_num, max_content_width + 2)

    st.download_button(
        label="Kombinierte Ergebnisse als Excel herunterladen",
        data=output.getvalue(),
        file_name="Kombinierte_Suchergebnisse_nach_KW.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
