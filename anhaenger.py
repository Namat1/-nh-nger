import streamlit as st
import pandas as pd
from io import BytesIO
import re

# Titel der App
st.title("Touren-Such-App")

# Datei-Upload
uploaded_file = st.file_uploader("Lade deine Excel- oder CSV-Datei hoch", type=["xlsx", "xls", "csv"])

if uploaded_file:
    try:
        # Dateiname extrahieren
        file_name = uploaded_file.name
        st.write(f"Hochgeladene Datei: {file_name}")

        # Kalenderwoche aus dem Dateinamen extrahieren
        kw_match = re.search(r'KW(\d{1,2})', file_name, re.IGNORECASE)
        kalenderwoche = f"KW{kw_match.group(1)}" if kw_match else "Keine KW gefunden"

        # Datei lesen
        if uploaded_file.name.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(uploaded_file, sheet_name="Touren")
            st.success("Das Blatt 'Touren' wurde erfolgreich geladen!")
        else:
            df = pd.read_csv(uploaded_file)
            st.success("CSV-Datei wurde erfolgreich geladen!")

        # Ursprüngliche Daten anzeigen
        st.write("Originaldaten:")
        st.dataframe(df)

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
                'Unnamed: 12': 'Art',
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

            # Zusammenfassung erstellen
            summary = final_results.groupby(['Nachname', 'Vorname'])['Verdienst'].sum().reset_index()
            summary = summary.rename(columns={"Verdienst": "Gesamtverdienst"})

            # Ergebnisse anzeigen
            st.write("Suchergebnisse:")
            st.dataframe(final_results)

            st.write("Zusammenfassung:")
            st.dataframe(summary)

            # Ergebnisse in Excel exportieren
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Suchergebnisse
                workbook = writer.book
                worksheet = workbook.add_worksheet("Suchergebnisse")
                writer.sheets["Suchergebnisse"] = worksheet
                worksheet.write(0, 0, f"Kalenderwoche: {kalenderwoche}")
                final_results.to_excel(writer, index=False, sheet_name="Suchergebnisse", startrow=2)
                for col_idx, column_name in enumerate(final_results.columns):
                    col_width = max(final_results[column_name].astype(str).map(len).max(), len(column_name))
                    worksheet.set_column(col_idx, col_idx, col_width)

                # Zusammenfassung
                summary_worksheet = workbook.add_worksheet("Zusammenfassung")
                writer.sheets["Zusammenfassung"] = summary_worksheet
                summary.to_excel(writer, index=False, sheet_name="Zusammenfassung", startrow=0)
                for col_idx, column_name in enumerate(summary.columns):
                    col_width = max(summary[column_name].astype(str).map(len).max(), len(column_name))
                    summary_worksheet.set_column(col_idx, col_idx, col_width)

            # Download-Button
            st.download_button(
                label="Suchergebnisse und Zusammenfassung als Excel herunterladen",
                data=output.getvalue(),
                file_name="Suchergebnisse_mit_Zusammenfassung.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            missing_columns = [col for col in required_columns if col not in df.columns]
            st.error(f"Die folgenden Spalten fehlen: {', '.join(missing_columns)}")

    except Exception as e:
        st.error(f"Fehler beim Verarbeiten der Datei: {e}")
else:
    st.info("Bitte lade eine Excel- oder CSV-Datei hoch, um zu starten.")
