import streamlit as st
import pandas as pd
from io import BytesIO
import re
import time

# Titel der App
st.title("Touren-Such-App für mehrere Dateien mit Zusammenfassung nach KW")

# Mehrere Dateien hochladen
uploaded_files = st.file_uploader("Lade deine Excel- oder CSV-Dateien hoch", type=["xlsx", "xls", "csv"], accept_multiple_files=True)

if uploaded_files:
    # Fortschrittsbalken erstellen
    progress_bar = st.progress(0)  # Initialisiere den Fortschrittsbalken
    total_files = len(uploaded_files)
    processed_files = 0

    all_results = []  # Liste, um Ergebnisse zu speichern
    all_summaries = []  # Liste, um Zusammenfassungen zu speichern

    for uploaded_file in uploaded_files:
        try:
            # Fortschrittsanzeige aktualisieren
            processed_files += 1
            progress_percentage = int((processed_files / total_files) * 100)
            progress_bar.progress(progress_percentage)

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

                # Zeilen mit 0 oder NaN in "Verdienst" entfernen (Numerischer Vergleich)
                final_results = final_results[(final_results['Verdienst'] > 0) & final_results['Verdienst'].notna()]

                # Euro-Zeichen in den Suchergebnissen hinzufügen
                final_results['Verdienst'] = final_results['Verdienst'].apply(lambda x: f"{x} €")

                # KW zur Ergebnis-Tabelle hinzufügen
                final_results['KW'] = kalenderwoche

                # Ergebnisse sammeln
                all_results.append(final_results)

                # Zusammenfassung erstellen (numerisch summieren)
                summary = final_results.copy()
                summary['Verdienst'] = summary['Verdienst'].str.replace(" €", "", regex=False).astype(float)  # Entferne Euro-Zeichen
                summary = summary.groupby(['KW', 'Nachname', 'Vorname']).agg({'Verdienst': 'sum'}).reset_index()

                # Euro-Zeichen hinzufügen in der Zusammenfassung und Spalte umbenennen
                summary['Gesamtverdienst'] = summary['Verdienst'].apply(lambda x: f"{x} €")
                summary = summary.drop(columns=['Verdienst'])  # Spalte 'Verdienst' entfernen

                # Zusammenfassung in die Sammlung einfügen
                all_summaries.append(summary)
        except Exception:
            pass  # Keine Fehlermeldungen anzeigen

    # Gesamtergebnisse zusammenführen
    if all_results:
        combined_results = pd.concat(all_results, ignore_index=True)

        # Entfernen unerwünschter Spalten
        columns_to_drop = [col for col in ['Datei', 'Art'] if col in combined_results.columns]
        final_output_results = combined_results.drop(columns=columns_to_drop)

        combined_summary = pd.concat(all_summaries, ignore_index=True)

        # Ergebnisse in eine Excel-Datei exportieren
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Suchergebnisse
            worksheet = writer.book.add_worksheet("Suchergebnisse")
            final_output_results.to_excel(writer, index=False, sheet_name="Suchergebnisse", startrow=2)

            # Auto-Spaltenbreite für Suchergebnisse
            for col_idx, column_name in enumerate(final_output_results.columns):
                max_content_width = max(final_output_results[column_name].astype(str).map(len).max(), len(column_name))
                worksheet.set_column(col_idx, col_idx, max_content_width + 2)

            # Zusammenfassung nach KW
            summary_worksheet = writer.book.add_worksheet("Zusammenfassung")
            combined_summary.to_excel(writer, index=False, sheet_name="Zusammenfassung", startrow=2)

            # Auto-Spaltenbreite für Zusammenfassung
            for col_idx, column_name in enumerate(combined_summary.columns):
                max_content_width = max(combined_summary[column_name].astype(str).map(len).max(), len(column_name))
                summary_worksheet.set_column(col_idx, col_idx, max_content_width + 2)

        # Download-Button
        st.download_button(
            label="Kombinierte Ergebnisse als Excel herunterladen",
            data=output.getvalue(),
            file_name="Kombinierte_Suchergebnisse_nach_KW.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
