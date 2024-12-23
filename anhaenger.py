import streamlit as st
import pandas as pd
from io import BytesIO
import re

# Titel der App
st.title("Touren-Such-App für mehrere Dateien mit Zusammenfassung nach KW")

# Mehrere Dateien hochladen
uploaded_files = st.file_uploader("Lade deine Excel- oder CSV-Dateien hoch", type=["xlsx", "xls", "csv"], accept_multiple_files=True)

if uploaded_files:
    all_results = []  # Liste, um Ergebnisse zu speichern
    all_summaries = []  # Liste, um Zusammenfassungen zu speichern

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

                final_results['Verdienst'] = final_results.apply(calculate_payment, axis=1)
                final_results['KW'] = kalenderwoche  # KW zur Ergebnis-Tabelle hinzufügen

                # Zeilen mit 0 oder NaN in "Verdienst" entfernen
                final_results = final_results[(final_results['Verdienst'] > 0) & final_results['Verdienst'].notna()]

                # Ergebnisse sammeln
                all_results.append(final_results)
                summary = final_results.groupby(['KW', 'Nachname', 'Vorname'])['Verdienst'].sum().reset_index()
                summary = summary.rename(columns={"Verdienst": "Gesamtverdienst"})
                all_summaries.append(summary)
            else:
                missing_columns = [col for col in required_columns if col not in df.columns]
                st.error(f"Die Datei {file_name} fehlt folgende Spalten: {', '.join(missing_columns)}")

        except Exception as e:
            st.error(f"Fehler beim Verarbeiten der Datei {file_name}: {e}")

    # Gesamtergebnisse zusammenführen
    if all_results:
        combined_results = pd.concat(all_results, ignore_index=True)

        # Entfernen unerwünschter Spalten
        columns_to_drop = [col for col in ['Datei', 'Art'] if col in combined_results.columns]
        final_output_results = combined_results.drop(columns=columns_to_drop)

        combined_summary = pd.concat(all_summaries, ignore_index=True)

        # Gesamte Zusammenfassung anzeigen
        st.write("Kombinierte Suchergebnisse:")
        st.dataframe(final_output_results)
        st.write("Zusammenfassung nach KW:")
        st.dataframe(combined_summary)

                # Ergebnisse in eine Excel-Datei exportieren
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Suchergebnisse
            workbook = writer.book
            worksheet = workbook.add_worksheet("Suchergebnisse")
            writer.sheets["Suchergebnisse"] = worksheet

            # Daten schreiben
            final_output_results.to_excel(writer, index=False, sheet_name="Suchergebnisse", startrow=2)

            # Verfeinerte Zellenbreitenanpassung
            for col_idx, column_name in enumerate(final_output_results.columns):
                max_content_width = final_output_results[column_name].astype(str).map(len).max()
                max_header_width = len(column_name)
                adjusted_width = max(max_content_width, max_header_width) + 2  # Puffer von 2 Zeichen
                worksheet.set_column(col_idx, col_idx, adjusted_width)

            # Zusammenfassung nach KW
            for kw in combined_summary['KW'].unique():
                summary_by_kw = combined_summary[combined_summary['KW'] == kw]
                summary_worksheet = workbook.add_worksheet(f"Zusammenfassung_{kw}")
                writer.sheets[f"Zusammenfassung_{kw}"] = summary_worksheet

                # Schreiben der zusammengefassten Daten für jede KW
                summary_by_kw.to_excel(writer, index=False, sheet_name=f"Zusammenfassung_{kw}", startrow=0)

                # Verfeinerte Zellenbreitenanpassung für Zusammenfassung
                for col_idx, column_name in enumerate(summary_by_kw.columns):
                    max_content_width = summary_by_kw[column_name].astype(str).map(len).max()
                    max_header_width = len(column_name)
                    adjusted_width = max(max_content_width, max_header_width) + 2  # Puffer von 2 Zeichen
                    summary_worksheet.set_column(col_idx, col_idx, adjusted_width)

        # Download-Button
        st.download_button(
            label="Kombinierte Ergebnisse als Excel herunterladen",
            data=output.getvalue(),
            file_name="Kombinierte_Suchergebnisse_nach_KW.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
