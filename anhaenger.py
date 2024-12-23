import streamlit as st
import pandas as pd
from io import BytesIO
import re

# Titel der App
st.title("Touren-Such-App für mehrere Dateien mit Zusammenfassung nach KW")

# Mehrere Dateien hochladen
uploaded_files = st.file_uploader("Lade deine Excel- oder CSV-Dateien hoch", type=["xlsx", "xls", "csv"], accept_multiple_files=True)

if uploaded_files:
    progress_bar = st.progress(0)
    total_files = len(uploaded_files)
    processed_files = 0

    all_results = []
    all_summaries = []

    for uploaded_file in uploaded_files:
        try:
            processed_files += 1
            progress_percentage = int((processed_files / total_files) * 100)
            progress_bar.progress(progress_percentage)

            file_name = uploaded_file.name
            kw_match = re.search(r'KW(\d{1,2})', file_name, re.IGNORECASE)
            kalenderwoche = f"KW{kw_match.group(1)}" if kw_match else "Keine KW gefunden"

            if uploaded_file.name.endswith(('.xlsx', '.xls')):
                df = pd.read_excel(uploaded_file, sheet_name="Touren")
            else:
                df = pd.read_csv(uploaded_file)

            search_numbers = ["602", "620", "350", "520", "156"]
            search_strings = ["AZ", "Az", "az", "MW", "Mw", "mw"]
            required_columns = ['Unnamed: 0', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 6',
                                'Unnamed: 7', 'Unnamed: 11', 'Unnamed: 12', 'Unnamed: 14']

            if all(col in df.columns for col in required_columns):
                df['Unnamed: 11'] = df['Unnamed: 11'].astype(str)
                df['Unnamed: 14'] = df['Unnamed: 14'].astype(str)

                number_matches = df[df['Unnamed: 11'].isin(search_numbers) & (df['Unnamed: 11'] != "607")]
                text_matches = df[df['Unnamed: 14'].str.contains('|'.join(search_strings), case=False, na=False) & (df['Unnamed: 11'] != "607")]
                combined_results = pd.concat([number_matches, text_matches]).drop_duplicates()

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

        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            final_output_results.to_excel(writer, index=False, sheet_name="Suchergebnisse", startrow=2)
            summary_worksheet = writer.book.add_worksheet("Zusammenfassung")
            combined_summary.to_excel(writer, index=False, sheet_name="Zusammenfassung", startrow=1)

        st.download_button(
            label="Kombinierte Ergebnisse als Excel herunterladen",
            data=output.getvalue(),
            file_name="Kombinierte_Suchergebnisse_nach_KW.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
