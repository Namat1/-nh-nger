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

    # Fortschrittsanzeige schließen und "FERTIG" anzeigen
    progress_bar.empty()  # Fortschrittsbalken entfernen
    st.success("FERTIG! Alle Dateien wurden verarbeitet.")

# Download-Bereich
if combined_results is not None and combined_summary is not None:
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Suchergebnisse
        worksheet = writer.book.add_worksheet("Suchergebnisse")
        combined_results.to_excel(writer, index=False, sheet_name="Suchergebnisse", startrow=2)

        # Auto-Spaltenbreite für Suchergebnisse
        for col_idx, column_name in enumerate(combined_results.columns):
            max_content_width = max(combined_results[column_name].astype(str).map(len).max(), len(column_name))
            worksheet.set_column(col_idx, col_idx, max_content_width + 2)

        # Zusammenfassung nach KW
        summary_worksheet = writer.book.add_worksheet("Zusammenfassung")
        combined_summary.to_excel(writer, index=False, sheet_name="Zusammenfassung", startrow=1)

        # Formatierungen hinzufügen
        header_format = writer.book.add_format({'bold': True, 'bg_color': '#b1c588', 'border': 1})
        blue_format = writer.book.add_format({'bg_color': '#67aadb', 'border': 1})
        green_format = writer.book.add_format({'bg_color': '#8cba8f', 'border': 1})

        # Formatierung der Kopfzeile
        for col_idx, column_name in enumerate(combined_summary.columns):
            summary_worksheet.write(0, col_idx, column_name, header_format)

        # Zeilen formatieren mit Trennung nach KW
        current_kw = None
        current_format = green_format
        for row_idx in range(len(combined_summary)):
            kw = combined_summary.iloc[row_idx, 0]  # KW-Wert
            if kw != current_kw:
                current_kw = kw
                current_format = green_format if current_format == blue_format else blue_format

            # Zellen formatieren
            for col_idx in range(len(combined_summary.columns)):
                summary_worksheet.write(row_idx + 1, col_idx, combined_summary.iloc[row_idx, col_idx], current_format)

        # Auto-Spaltenbreite für Zusammenfassung
        for col_idx, column_name in enumerate(combined_summary.columns):
            max_content_width = max(combined_summary[column_name].astype(str).map(len).max(), len(column_name))
            summary_worksheet.set_column(col_idx, col_idx, max_content_width + 2)

    st.download_button(
        label="Kombinierte Ergebnisse als Excel herunterladen",
        data=output.getvalue(),
        file_name="Kombinierte_Suchergebnisse_nach_KW.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
