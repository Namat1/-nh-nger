import streamlit as st
import pandas as pd
from io import BytesIO
import re

# Titel der App
st.title("Touren-Such-App für mehrere Dateien mit Zusammenfassung nach KW")

# Mehrere Dateien hochladen
uploaded_files = st.file_uploader("Lade deine Excel- oder CSV-Dateien hoch", type=["xlsx", "xls", "csv"], accept_multiple_files=True)

if uploaded_files:
    # Fortschrittsbalken erstellen
    progress_bar = st.progress(0)
    total_files = len(uploaded_files)
    processed_files = 0

    all_results = []
    all_summaries = []

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
                df['Unnamed: 11'] = df['Unnamed: 11'].astype(str)
                df['Unnamed: 14'] = df['Unnamed: 14'].astype(str)

                # Filter nach Suchoptionen
                number_matches = df[df['Unnamed: 11'].isin(search_numbers) & (df['Unnamed: 11'] != "607")]
                text_matches = df[df['Unnamed: 14'].str.contains('|'.join(search_strings), case=False, na=False) & (df['Unnamed: 11'] != "607")]
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
                final_results = final_results.sort_values(by=['Nachname', 'Vorname'])

                # Verdienst berechnen
                payment_mapping = {"602": 40, "156": 40, "620": 20, "350": 20, "520": 20}

                def calculate_payment(row):
                    kennzeichen = row['Kennzeichen']
                    art_2 = row['Art 2'].strip().upper()
                    return payment_mapping.get(kennzeichen, 0) if art_2 == "AZ" else 0

                final_results['Verdienst'] = final_results.apply(calculate_payment, axis=1)
                final_results = final_results[(final_results['Verdienst'] > 0) & final_results['Verdienst'].notna()]
                final_results['Verdienst'] = final_results['Verdienst'].apply(lambda x: f"{x} €")
                final_results['KW'] = kalenderwoche
                all_results.append(final_results)

                # Zusammenfassung erstellen
                summary = final_results.copy()
                summary['Verdienst'] = summary['Verdienst'].str.replace(" €", "", regex=False).astype(float)
                summary = summary.groupby(['KW', 'Nachname', 'Vorname']).agg({'Verdienst': 'sum'}).reset_index()
                summary['Gesamtverdienst'] = summary['Verdienst'].apply(lambda x: f"{x} €")
                summary = summary.drop(columns=['Verdienst'])
                all_summaries.append(summary)
        except Exception:
            pass

    if all_results:
        combined_results = pd.concat(all_results, ignore_index=True)
        columns_to_drop = [col for col in ['Datei', 'Art'] if col in combined_results.columns]
        final_output_results = combined_results.drop(columns=columns_to_drop)
        combined_summary = pd.concat(all_summaries, ignore_index=True)

        # Ergebnisse in eine Excel-Datei exportieren
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Suchergebnisse exportieren
            final_output_results.to_excel(writer, index=False, sheet_name="Suchergebnisse", startrow=2)
            worksheet = writer.sheets["Suchergebnisse"]

            # Auto-Spaltenbreite für Suchergebnisse
            for col_idx, column_name in enumerate(final_output_results.columns):
                max_content_width = max(final_output_results[column_name].astype(str).map(len).max(), len(column_name))
                worksheet.set_column(col_idx, col_idx, max_content_width + 2)

            # Zusammenfassung exportieren
            combined_summary.to_excel(writer, index=False, sheet_name="Zusammenfassung", startrow=1)
            summary_worksheet = writer.sheets["Zusammenfassung"]

            # Farben und Kopfzeilenformat hinzufügen
            header_format = writer.book.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
            blue_format = writer.book.add_format({'bg_color': '#E3F2FD', 'border': 1})
            green_format = writer.book.add_format({'bg_color': '#E8F5E9', 'border': 1})

            # Kopfzeilen formatieren
            for col_idx, column_name in enumerate(combined_summary.columns):
                summary_worksheet.write(0, col_idx, column_name, header_format)

            # Zeilen mit Farben für Kalenderwoche
            current_kw = None
            current_format = green_format
            for row_idx in range(len(combined_summary)):
                kw = combined_summary.iloc[row_idx, 0]
                if kw != current_kw:
                    current_kw = kw
                    current_format = green_format if current_format == blue_format else blue_format
                for col_idx in range(len(combined_summary.columns)):
                    summary_worksheet.write(row_idx + 1, col_idx, combined_summary.iloc[row_idx, col_idx], current_format)

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
