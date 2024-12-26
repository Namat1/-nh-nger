import streamlit as st
import pandas as pd
from io import BytesIO
import re

# Titel der App
st.title("Zulage Sonderfahrzeuge")

# Mehrere Dateien hochladen
uploaded_files = st.file_uploader("Lade deine Excel- oder CSV-Dateien hoch", type=["xlsx", "xls", "csv"], accept_multiple_files=True)

# Variablen zur Zwischenspeicherung
combined_results = None

if uploaded_files:
    all_results = []

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
                df = pd.read_excel(uploaded_file, sheet_name="Touren", dtype=str)
            else:
                df = pd.read_csv(uploaded_file, dtype=str)

            # Sicherstellen, dass alle benötigten Spalten vorhanden sind
            if 'Unnamed: 6' not in df.columns:
                df['Unnamed: 6'] = ""  # Leere Spalte hinzufügen
            if 'Unnamed: 7' not in df.columns:
                df['Unnamed: 7'] = ""  # Leere Spalte hinzufügen

            # Filteroptionen
            search_numbers = ["602", "620", "350", "520", "156"]
            search_strings = ["AZ", "Az", "az", "MW", "Mw", "mw"]

            required_columns = ['Unnamed: 0', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 6',
                                'Unnamed: 7', 'Unnamed: 11', 'Unnamed: 14']

            if all(col in df.columns for col in required_columns):
                # Umbenennen der Spalten für Klarheit
                df.rename(
                    columns={
                        'Unnamed: 0': 'Tour',
                        'Unnamed: 3': 'Nachname',
                        'Unnamed: 4': 'Vorname',
                        'Unnamed: 6': 'Nachname 2',
                        'Unnamed: 7': 'Vorname 2',
                        'Unnamed: 11': 'Kennzeichen',
                        'Unnamed: 14': 'Art 2'
                    },
                    inplace=True
                )

                # Filtern der Daten
                number_matches = df[df['Kennzeichen'].isin(search_numbers)]
                text_matches = df[df['Art 2'].str.contains('|'.join(search_strings), case=False, na=False)]
                combined_results_df = pd.concat([number_matches, text_matches]).drop_duplicates()

                # Kategorie hinzufügen
                combined_results_df['Kategorie'] = combined_results_df['Kennzeichen'].map(
                    lambda x: "Gruppe 1 (156, 602)" if x in ["156", "602"] else
                              "Gruppe 2 (620, 350, 520)" if x in ["620", "350", "520"] else "Andere"
                )

                # Kalenderwoche hinzufügen
                combined_results_df['KW'] = kalenderwoche
                all_results.append(combined_results_df)

        except Exception as e:
            st.error(f"Fehler beim Verarbeiten der Datei {file_name}: {e}")

    if all_results:
        combined_results = pd.concat(all_results, ignore_index=True).fillna("")

    progress_bar.empty()
    st.success("FERTIG! Alle Dateien wurden verarbeitet.")

if combined_results is not None and not combined_results.empty:

    # Pivot-Tabelle für Blatt 3 erstellen
    vehicle_grouped = combined_results.pivot_table(
        index=['Kategorie', 'KW', 'Nachname', 'Vorname', 'Nachname 2', 'Vorname 2'],
        columns='Kennzeichen',
        values='Art 2',
        aggfunc='count',
        fill_value=0
    ).reset_index()

    # Gesamtsumme berechnen
    vehicle_grouped['Gesamtsumme'] = vehicle_grouped.iloc[:, 6:].sum(axis=1)

    # Export der Daten in eine Excel-Datei
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Blatt 3: Auflistung Fahrzeuge
        vehicle_grouped.to_excel(writer, sheet_name="Auflistung Fahrzeuge", index=False)

        # Formatierung der Spaltenbreiten
        worksheet = writer.sheets["Auflistung Fahrzeuge"]
        for col_num, column_name in enumerate(vehicle_grouped.columns):
            max_width = max(vehicle_grouped[column_name].astype(str).map(len).max(), len(column_name), 10)
            worksheet.set_column(col_num, col_num, max_width + 2)

    output.seek(0)
    st.download_button(
        label="Ergebnisse als Excel herunterladen",
        data=output,
        file_name="Ergebnisse_Auflistung_Fahrzeuge.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
