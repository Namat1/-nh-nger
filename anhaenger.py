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

        # Sortierung der Suchergebnisse nach numerischer KW
        combined_results['KW_Numeric'] = combined_results['KW'].str.extract(r'(\d+)').astype(int)
        combined_results = combined_results.sort_values(by=['KW_Numeric', 'Nachname', 'Vorname']).drop(columns=['KW_Numeric'])

        # Sortierung der Zusammenfassung nach numerischer KW
        combined_summary = pd.concat(all_summaries, ignore_index=True)
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
        combined_results.to_excel(writer, index=False, sheet_name="Suchergebnisse")

        # Formatierungen hinzufügen
        worksheet = writer.sheets['Suchergebnisse']

        # Spaltenbreite automatisch anpassen
        for col_num, column_cells in enumerate(worksheet.iter_cols(min_row=1, max_col=worksheet.max_column, max_row=worksheet.max_row), start=1):
            max_length = 0
            for cell in column_cells:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            adjusted_width = max_length + 2
            worksheet.set_column(col_num - 1, col_num - 1, adjusted_width)

       # Blatt 3: Fahrzeuggruppen
combined_results['Kategorie'] = combined_results['Kennzeichen'].map(
    lambda x: "Gruppe 1 (156, 602)" if x in ["156", "602"] else
              "Gruppe 2 (620, 350, 520)" if x in ["620", "350", "520"] else "Andere"
)
vehicle_grouped = combined_results.pivot_table(
    index=['Kategorie', 'KW', 'Nachname', 'Vorname'],
    columns='Kennzeichen',
    values='Verdienst',
    aggfunc='sum',
    fill_value=0
).reset_index()

# Sicherstellen, dass die Fahrzeugspalten numerisch sind
vehicle_grouped.iloc[:, 4:] = vehicle_grouped.iloc[:, 4:].apply(pd.to_numeric, errors='coerce')

# KW numerisch sortieren
vehicle_grouped = vehicle_grouped.sort_values(by='KW')

# Summenspalte hinzufügen
vehicle_grouped['Gesamtsumme (€)'] = vehicle_grouped.iloc[:, 4:].sum(axis=1)

# Formatierung für Euro
for col in vehicle_grouped.columns[4:]:
    vehicle_grouped[col] = vehicle_grouped[col].apply(lambda x: f"{x:.2f} €")

# Daten in Excel schreiben
vehicle_grouped.to_excel(writer, sheet_name="Fahrzeuggruppen", index=False)
worksheet = writer.sheets["Fahrzeuggruppen"]

# Spaltenbreite automatisch anpassen
for column_cells in worksheet.columns:
    max_length = 0
    column = column_cells[0].column_letter  # Get the column name
    for cell in column_cells:
        try:  # Necessary to avoid issues with empty cells
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        except:
            pass
    adjusted_width = max_length + 2
    worksheet.column_dimensions[column].width = adjusted_width

# Farbige Hervorhebung nach KW
kw_colors = {
    range(0, 100): 'FFB3E5FC',  # Hellblau
    range(100, 200): 'FFC8E6C9',  # Hellgrün
    range(200, 300): 'FFFFF9C4',  # Hellgelb
}

for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=2, max_col=2):
    kw_cell = row[0]
    kw_value = int(kw_cell.value)
    for kw_range, color in kw_colors.items():
        if kw_value in kw_range:
            kw_cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
            break


    # Download-Link erstellen
    st.download_button(
        label="Kombinierte Ergebnisse als Excel herunterladen",
        data=output.getvalue(),
        file_name="Kombinierte_Suchergebnisse_nach_KW.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
