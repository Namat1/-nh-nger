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
    all_results = []
    all_summaries = []

    progress_bar = st.progress(0)
    total_files = len(uploaded_files)

    for idx, uploaded_file in enumerate(uploaded_files):
        try:
            progress_bar.progress((idx + 1) / total_files)
            file_name = uploaded_file.name

            # Kalenderwoche aus Dateinamen extrahieren
            kw_match = re.search(r'KW(\d{1,2})', file_name, re.IGNORECASE)
            kalenderwoche = f"KW{kw_match.group(1)}" if kw_match else "Keine KW gefunden"

            # Datei lesen
            if uploaded_file.name.endswith(('.xlsx', '.xls')):
                df = pd.read_excel(uploaded_file, sheet_name="Touren")
            else:
                df = pd.read_csv(uploaded_file)

            # Filteroptionen
            search_numbers = ["602", "620", "350", "520", "156"]
            search_strings = ["AZ", "Az", "az", "MW", "Mw", "mw"]

            required_columns = ['Unnamed: 0', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 6',
                                'Unnamed: 7', 'Unnamed: 11', 'Unnamed: 12', 'Unnamed: 14']

            if all(col in df.columns for col in required_columns):
                df['Unnamed: 11'] = df['Unnamed: 11'].astype(str)
                df['Unnamed: 14'] = df['Unnamed: 14'].astype(str)

                number_matches = df[df['Unnamed: 11'].isin(search_numbers) & (df['Unnamed: 11'] != "607")]
                text_matches = df[df['Unnamed: 14'].str.contains('|'.join(search_strings), case=False, na=False) & 
                                  (df['Unnamed: 11'] != "607")]
                combined_results_df = pd.concat([number_matches, text_matches]).drop_duplicates()

                # Spalten umbenennen
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
        except Exception as e:
            st.error(f"Fehler beim Verarbeiten der Datei {file_name}: {e}")

    if all_results:
        combined_results = pd.concat(all_results, ignore_index=True)
        combined_results['KW_Numeric'] = combined_results['KW'].str.extract(r'(\d+)').astype(int)
        combined_results = combined_results.sort_values(by=['KW_Numeric', 'Nachname', 'Vorname']).drop(columns=['KW_Numeric'])

        combined_summary = pd.concat(all_summaries, ignore_index=True)
        combined_summary['KW_Numeric'] = combined_summary['KW'].str.extract(r'(\d+)').astype(int)
        combined_summary = combined_summary.sort_values(by=['KW_Numeric', 'Nachname', 'Vorname']).drop(columns=['KW_Numeric'])

    progress_bar.empty()
    st.success("FERTIG! Alle Dateien wurden verarbeitet.")

if combined_results is not None and combined_summary is not None:
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Blatt 1: Suchergebnisse
        combined_results.to_excel(writer, index=False, sheet_name="Suchergebnisse")
        workbook = writer.book
        worksheet = writer.sheets['Suchergebnisse']

        unique_kws = combined_results['KW'].unique()
        colors = ["#FFEB9C", "#D9EAD3", "#F4CCCC", "#CFE2F3", "#FFD966"]
        formats = {kw: workbook.add_format({'bg_color': colors[i % len(colors)], 'border': 1}) for i, kw in enumerate(unique_kws)}
        default_format = workbook.add_format({'border': 1})

        for col_num, column_name in enumerate(combined_results.columns):
            max_width = max(combined_results[column_name].astype(str).map(len).max(), len(column_name), 10)
            worksheet.set_column(col_num, col_num, max_width + 2)

        for row_num, kw in enumerate(combined_results['KW'], start=1):
            row_format = formats.get(kw, default_format)
            worksheet.set_row(row_num, None, row_format)

        # Blatt 2: Zusammenfassung
        combined_summary.to_excel(writer, index=False, sheet_name="Auszahlung pro KW")
        summary_sheet = writer.sheets['Auszahlung pro KW']

        # Formatierungen hinzufügen
        header_format = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
        blue_format = workbook.add_format({'bg_color': '#76bef5', 'border': 1})
        green_format = workbook.add_format({'bg_color': '#6bff77', 'border': 1})

        for col_num, column_name in enumerate(combined_summary.columns):
            max_width = max(combined_summary[column_name].astype(str).apply(len).max(), len(column_name), 10)
            summary_sheet.set_column(col_num, col_num, max_width + 2)
            summary_sheet.write(0, col_num, column_name, header_format)

        current_kw = None
        current_format = green_format
        for row_num in range(len(combined_summary)):
            kw = combined_summary.iloc[row_num]['KW']
            if kw != current_kw:
                current_kw = kw
                current_format = green_format if current_format == blue_format else blue_format

            for col_num in range(len(combined_summary.columns)):
                summary_sheet.write(row_num + 1, col_num, combined_summary.iloc[row_num, col_num], current_format)

                                        # Fahrzeuggruppen kategorisieren
combined_results['Kategorie'] = combined_results['Kennzeichen'].map(
    lambda x: "Gruppe 1 (156, 602)" if x in ["156", "602"] else
              "Gruppe 2 (620, 350, 520)" if x in ["620", "350", "520"] else "Andere"
)

# Pivot-Tabelle erstellen
vehicle_grouped = combined_results.pivot_table(
    index=['Kategorie', 'KW', 'Nachname', 'Vorname'],
    columns='Kennzeichen',
    values='Verdienst',
    aggfunc=lambda x: sum(float(v.replace(" €", "")) for v in x if isinstance(v, str)),
    fill_value=0
).reset_index()

# Summenspalte hinzufügen
vehicle_grouped['Gesamtsumme (€)'] = vehicle_grouped.iloc[:, 4:].sum(axis=1)

# Formatierung der Summenwerte mit €
for col in vehicle_grouped.columns[4:]:
    vehicle_grouped[col] = vehicle_grouped[col].apply(lambda x: f"{x:.2f} €")

# Sortierung nach KW (numerisch)
vehicle_grouped['KW_Numeric'] = vehicle_grouped['KW'].str.extract(r'(\d+)').astype(int)
vehicle_grouped = vehicle_grouped.sort_values(by=['KW_Numeric', 'Kategorie', 'Nachname', 'Vorname'])
vehicle_grouped = vehicle_grouped.drop(columns=['KW_Numeric'])  # Sortierspalte entfernen

# Tabelle im Excel speichern
vehicle_grouped.to_excel(writer, sheet_name="Auflistung Fahrzeuge", index=False)
vehicle_sheet = writer.sheets['Auflistung Fahrzeuge']

# Dynamische Spaltenbreite
for col_num, column_name in enumerate(vehicle_grouped.columns):
    max_width = max(vehicle_grouped[column_name].astype(str).apply(len).max(), len(column_name), 10)
    vehicle_sheet.set_column(col_num, col_num, max_width + 2)

# Farben für die KW-Zeilen
kw_colors = ['#FFEB9C', '#D9EAD3', '#F4CCCC', '#CFE2F3', '#FFD966']
current_kw = None
current_color_index = 0

# Zeilen farblich nach KW formatieren
for row_num in range(len(vehicle_grouped)):
    kw = vehicle_grouped.iloc[row_num]['KW']
    if kw != current_kw:
        current_kw = kw
        current_color_index = (current_color_index + 1) % len(kw_colors)

    row_format = workbook.add_format({'bg_color': kw_colors[current_color_index], 'border': 1})

    for col_num, value in enumerate(vehicle_grouped.iloc[row_num]):
        vehicle_sheet.write(row_num + 1, col_num, value, row_format)

# Bold-Format für Kategorie und KW definieren
bold_format = workbook.add_format({'bold': True})

# Anwenden von Bold auf Kategorie- und KW-Spalte
for row_num in range(len(vehicle_grouped)):
    # Erste Spalte (Index 0) - Kategorie
    category = vehicle_grouped.iloc[row_num]['Kategorie']
    vehicle_sheet.write(row_num + 1, 0, category, bold_format)  # Spalte 0 fett formatieren

    # Zweite Spalte (Index 1) - KW
    kw = vehicle_grouped.iloc[row_num]['KW']
    vehicle_sheet.write(row_num + 1, 1, kw, bold_format)

# Streamlit Download-Button
output.seek(0)
st.download_button(
    label="Kombinierte Ergebnisse als Excel herunterladen",
    data=output.getvalue(),
    file_name="Kombinierte_Suchergebnisse_nach_KW.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
