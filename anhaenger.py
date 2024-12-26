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
combined_summary = None 

name_to_personalnummer = {
    "Adler": {"Philipp": "00041450"},
    "Auer": {"Frank": "00020795"},
    "Batkowski": {"Tilo": "00046601"},
    "Benabbes": {"Badr": "00048980"},
    "Biebow": {"Thomas": "00042004"},
    "Bläsing": {"Elmar": "00049093"},
    "Bursian": {"Ronny": "00025714"},
    "Buth": {"Sven": "00046673"},
    "Böhnke": {"Marcel": "00020833"},
    "Carstensen": {"Martin": "00042412"},
    "Chege": {"Moses Gichuru": "00046106"},
    "Dammasch": {"Bernd": "00019297"},
    "Demuth": {"Harry": "00020796"},
    "Doroszkiewicz": {"Bogumil": "00049132"},
    "Dürr": {"Holger": "00039164"},
    "Effenberger": {"Sven": "00030807"},
    "Engel": {"Raymond": "00033429"},
    "Fechner": {"Danny": "00043696", "Klaus": "00038278"},
    "Findeklee": {"Bernd": "00020804"},
    "Flint": {"Henryk": "00042414"},
    "Fuhlbrügge": {"Justin": "00046289"},
    "Gehrmann": {"Rayk": "00046702"},
    "Gheonea": {"Costel-Daniel": "00050877"},
    "Glanz": {"Björn": "00041914"},
    "Gnech": {"Torsten": "00018613"},
    "Greve": {"Nicole": "00040760"},
    "Guthmann": {"Fred": "00018328"},
    "Hagen": {"Andy": "00020271"},
    "Hartig": {"Sebastian": "00044120"},
    "Haus": {"David": "00046101"},
    "Heeser": {"Bernd": "00041916"},
    "Helm": {"Philipp": "00046685"},
    "Henkel": {"Bastian": "00048187"},
    "Holtz": {"Torsten": "00021159"},
    "Janikiewicz": {"Radoslaw": "00042159"},
    "Kelling": {"Jonas Ole": "00044140"},
    "Kleiber": {"Lutz": "00026255"},
    "Klemkow": {"Ralf": "00040634"},
    "Kollmann": {"Steffen": "00040988"},
    "König": {"Heiko": "00036341"},
    "Krazewski": {"Cezary": "00039463"},
    "Krieger": {"Christian": "00049092"},
    "Krull": {"Benjamin": "00044192"},
    "Lange": {"Michael": "00035407"},
    "Lewandowski": {"Kamil": "00041044"},
    "Likoonski": {"Vladimir": "00044766"},
    "Linke": {"Erich": "00048377"},
    "Lefkih": {"Houssni": "00052293"},
    "Ludolf": {"Michel": "00048814"},
    "Marouni": {"Ayyoub": "00048986"},
    "Mintel": {"Mario": "00046686"},
    "Ohlenroth": {"Nadja": "00042114"},
    "Ohms": {"Torsten": "00019300"},
    "Okoth": {"Tedy Omondi": "00046107"},
    "Oszmian": {"Jacub": "00039464"},
    "Pabst": {"Torsten": "00021976"},
    "Pawlak": {"Bartosz": "00036381"},
    "Piepke": {"Torsten": "00021390"},
    "Plinke": {"Killian": "00044137"},
    "Pogodski": {"Enrico": "00046668"},
    "Quint": {"Stefan": "00035718"},
    "Rimba": {"Rimba Gona": "00046108"},
    "Sarwatka": {"Heiko": "00028747"},
    "Scheil": {"Eric-Rene": "00038579", "Rene": "00020851"},
    "Schlichting": {"Michael": "00021452"},
    "Schlutt": {"Hubert": "00020880", "Rene": "00042932"},
    "Schmieder": {"Steffen": "00046286"},
    "Schneider": {"Matthias": "00045495"},
    "Schulz": {"Julian": "00049130", "Stephan": "00041558"},
    "Singh": {"Jagtar": "00040902"},
    "Stoltz": {"Thorben": "00040991"},
    "Thal": {"Jannic": "00046006"},
    "Tumanow": {"Vasilli": "00045019"},
    "Wachnowski": {"Klaus": "00026019"},
    "Wendel": {"Danilo": "00048994"},
    "Wille": {"Rene": "00021393"},
    "Wisniewski": {"Krzysztof": "00046550"},
    "Zander": {"Jan": "00042454"},
    "Zosel": {"Ingo": "00026303"},
}



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


                # Fehlende Werte und Leerzeichen behandeln
                final_results['Nachname'] = final_results['Nachname'].fillna("").str.strip()
                final_results['Vorname'] = final_results['Vorname'].fillna("").str.strip()
                final_results['Nachname 2'] = final_results['Nachname 2'].fillna("").str.strip()
                final_results['Vorname 2'] = final_results['Vorname 2'].fillna("").str.strip()
                # Fehlende Namen aus 'Nachname 2' und 'Vorname 2' ergänzen
                final_results['Nachname'] = final_results.apply(
                   lambda row: row['Nachname 2'] if pd.isna(row['Nachname']) or row['Nachname'] == '' else row['Nachname'],
                   axis=1
                )

                final_results['Vorname'] = final_results.apply(
                    lambda row: row['Vorname 2'] if pd.isna(row['Vorname']) or row['Vorname'] == '' else row['Vorname'],
                    axis=1
                )

                
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
                summary['Gesamtverdienst'] = summary['Verdienst'].apply(lambda x: f"{x:.2f} €")
                summary = summary.drop(columns=['Verdienst'])

                # Personalnummer hinzufügen
                summary['Personalnummer'] = summary.apply(
                    lambda row: name_to_personalnummer.get(
                        (row['Nachname'], row['Vorname']), "Unbekannt"
                    ),
                    axis=1
                )

                all_summaries.append(summary)
        except Exception as e:
            st.error(f"Fehler beim Verarbeiten der Datei {file_name}: {e}")

    if all_results:
        combined_results = pd.concat(all_results, ignore_index=True).fillna("")
        combined_summary = pd.concat(all_summaries, ignore_index=True).fillna("")

        # Sortieren der Daten
        combined_results['KW_Numeric'] = combined_results['KW'].str.extract(r'(\d+)').astype(float).fillna(-1).astype(int)
        combined_results = combined_results[combined_results['KW_Numeric'] != -1].sort_values(by=['KW_Numeric', 'Nachname', 'Vorname']).drop(columns=['KW_Numeric'])

        combined_summary['KW_Numeric'] = combined_summary['KW'].str.extract(r'(\d+)').astype(float).fillna(-1).astype(int)
        combined_summary = combined_summary[combined_summary['KW_Numeric'] != -1].sort_values(by=['KW_Numeric', 'Nachname', 'Vorname']).drop(columns=['KW_Numeric'])

    progress_bar.empty()
    st.success("FERTIG! Alle Dateien wurden verarbeitet.")

if combined_results is not None and not combined_results.empty and combined_summary is not None and not combined_summary.empty:

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        kw_colors = ['#FFEB9C', '#D9EAD3', '#F4CCCC', '#CFE2F3', '#FFD966']
        current_kw = None
        current_color_index = 0

        # Blatt 1: Suchergebnisse
        combined_results.to_excel(writer, index=False, sheet_name="Suchergebnisse")
        worksheet = writer.sheets['Suchergebnisse']
        worksheet.freeze_panes(1, 0)  # Fixiert die erste Zeile
        worksheet.autofilter(0, 0, len(combined_results), len(combined_results.columns) - 1)  # Filter hinzufügen
        for col_num, column_name in enumerate(combined_results.columns):
            max_width = max(combined_results[column_name].astype(str).map(len).max(), len(column_name), 10)
            worksheet.set_column(col_num, col_num, max_width + 2)

        # Farben anwenden
        for row_num in range(len(combined_results)):
            kw = combined_results.iloc[row_num]['KW']
            if kw != current_kw:
                current_kw = kw
                current_color_index = (current_color_index + 1) % len(kw_colors)
            row_format = workbook.add_format({'bg_color': kw_colors[current_color_index], 'border': 1})
            for col_num, value in enumerate(combined_results.iloc[row_num]):
                worksheet.write(row_num + 1, col_num, str(value), row_format)

        # Blatt 2: Auszahlung pro KW
        combined_summary['Personalnummer'] = combined_summary.apply(
    lambda row: name_to_personalnummer.get(
        row['Nachname'], {}
    ).get(row['Vorname'], "Unbekannt"),
    axis=1
)

        combined_summary.to_excel(writer, index=False, sheet_name="Auszahlung pro KW")
        summary_sheet = writer.sheets['Auszahlung pro KW']
        summary_sheet.freeze_panes(1, 0)  # Fixiert die erste Zeile
        summary_sheet.autofilter(0, 0, len(combined_summary), len(combined_summary.columns) - 1)  # Filter hinzufügen
        for col_num, column_name in enumerate(combined_summary.columns):
            max_width = max(combined_summary[column_name].astype(str).map(len).max(), len(column_name), 10)
            summary_sheet.set_column(col_num, col_num, max_width + 2)


        for row_num in range(len(combined_summary)):
            kw = combined_summary.iloc[row_num]['KW']
            if kw != current_kw:
                current_kw = kw
                current_color_index = (current_color_index + 1) % len(kw_colors)
            row_format = workbook.add_format({'bg_color': kw_colors[current_color_index], 'border': 1})
            for col_num, value in enumerate(combined_summary.iloc[row_num]):
                summary_sheet.write(row_num + 1, col_num, str(value), row_format)

        # Blatt 3: Auflistung Fahrzeuge
        combined_results['Kategorie'] = combined_results['Kennzeichen'].map(
            lambda x: "Gruppe 1 (156, 602)" if x in ["156", "602"] else
                      "Gruppe 2 (620, 350, 520)" if x in ["620", "350", "520"] else "Andere"
        )
        vehicle_grouped = combined_results.pivot_table(
            index=['Kategorie', 'KW', 'Nachname', 'Vorname'],
            columns='Kennzeichen',
            values='Verdienst',
            aggfunc=lambda x: sum(float(v.replace(" €", "")) for v in x if isinstance(v, str)),
            fill_value=0
        ).reset_index()

        vehicle_grouped['Gesamtsumme (€)'] = vehicle_grouped.iloc[:, 4:].sum(axis=1)
        for col in vehicle_grouped.columns[4:]:
            vehicle_grouped[col] = vehicle_grouped[col].apply(lambda x: f"{x:.2f} €")

        vehicle_grouped['KW_Numeric'] = vehicle_grouped['KW'].str.extract(r'(\d+)').astype(int)
        vehicle_grouped = vehicle_grouped.sort_values(by=['KW_Numeric', 'Kategorie', 'Nachname', 'Vorname']).drop(columns=['KW_Numeric'])

        vehicle_grouped.to_excel(writer, sheet_name="Auflistung Fahrzeuge", index=False)
        vehicle_sheet = writer.sheets['Auflistung Fahrzeuge']
        vehicle_sheet.freeze_panes(1, 0)  # Fixiert die erste Zeile
        vehicle_sheet.autofilter(0, 0, len(vehicle_grouped), len(vehicle_grouped.columns) - 1)  # Filter hinzufügen
        for col_num, column_name in enumerate(vehicle_grouped.columns):
            max_width = max(vehicle_grouped[column_name].astype(str).map(len).max(), len(column_name), 10)
            vehicle_sheet.set_column(col_num, col_num, max_width + 2)

        for row_num in range(len(vehicle_grouped)):
            kw = vehicle_grouped.iloc[row_num]['KW']
            if kw != current_kw:
                current_kw = kw
                current_color_index = (current_color_index + 1) % len(kw_colors)
            row_format = workbook.add_format({'bg_color': kw_colors[current_color_index], 'border': 1})
            for col_num, value in enumerate(vehicle_grouped.iloc[row_num]):
                vehicle_sheet.write(row_num + 1, col_num, str(value), row_format)

    output.seek(0)
    st.download_button(
        label="Kombinierte Ergebnisse als Excel herunterladen",
        data=output.getvalue(),
        file_name="Kombinierte_Suchergebnisse_nach_KW.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
