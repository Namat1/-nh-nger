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
    all_summaries = []  # Liste, um Zusammenfassungen zu speichern

    progress_bar = st.progress(0)  # Fortschrittsbalken hinzufügen
    total_files = len(uploaded_files)

    for idx, uploaded_file in enumerate(uploaded_files):
        try:
            progress_bar.progress((idx + 1) / total_files)
            file_name = uploaded_file.name

            kw_match = re.search(r'KW(\d{1,2})', file_name, re.IGNORECASE)
            kalenderwoche = f"KW{kw_match.group(1)}" if kw_match else "Keine KW gefunden"

            if uploaded_file.name.endswith(('.xlsx', '.xls')):
                df = pd.read_excel(uploaded_file, sheet_name="Touren")
            else:
                df = pd.read_csv(uploaded_file)

            search_numbers = ["602", "620", "350", "520", "156"]
            required_columns = ['Unnamed: 0', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 6',
                                'Unnamed: 7', 'Unnamed: 11', 'Unnamed: 12', 'Unnamed: 14']

            if all(col in df.columns for col in required_columns):
                df['Unnamed: 11'] = df['Unnamed: 11'].astype(str)
                number_matches = df[df['Unnamed: 11'].isin(search_numbers)]
                combined_results_df = number_matches

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

                payment_mapping = {"602": 40, "156": 40, "620": 20, "350": 20, "520": 20}
                final_results['Verdienst'] = final_results['Kennzeichen'].map(payment_mapping).fillna(0)
                final_results['KW'] = kalenderwoche
                all_results.append(final_results)
        except Exception as e:
            st.error(f"Fehler beim Verarbeiten der Datei {file_name}: {e}")

    if all_results:
        combined_results = pd.concat(all_results, ignore_index=True)
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Blatt 1: Suchergebnisse
            combined_results.to_excel(writer, index=False, sheet_name="Suchergebnisse")

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

            # Summenspalte hinzufügen
            vehicle_grouped['Gesamtsumme (€)'] = vehicle_grouped.iloc[:, 4:].sum(axis=1)

            # Formatierung für Euro
            for col in vehicle_grouped.columns[4:]:
                vehicle_grouped[col] = vehicle_grouped[col].apply(lambda x: f"{x:.2f} €")
            vehicle_grouped.to_excel(writer, sheet_name="Fahrzeuggruppen", index=False)

        st.download_button(
            label="Kombinierte Ergebnisse als Excel herunterladen",
            data=output.getvalue(),
            file_name="Kombinierte_Suchergebnisse_nach_KW.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
