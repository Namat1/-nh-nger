import streamlit as st
import pandas as pd
from io import BytesIO
import re
import xlsxwriter

# Titel der App
st.title("Touren-Such-App")

# Datei-Upload
uploaded_file = st.file_uploader("Lade deine Excel- oder CSV-Datei hoch", type=["xlsx", "xls", "csv"])

if uploaded_file:
    try:
        # Extrahiere den Dateinamen
        file_name = uploaded_file.name
        st.write(f"Hochgeladene Datei: {file_name}")

        # Kalenderwoche aus dem Dateinamen extrahieren
        kw_match = re.search(r'KW(\d{1,2})', file_name, re.IGNORECASE)
        kalenderwoche = f"KW{kw_match.group(1)}" if kw_match else "Keine KW gefunden"

        # Prüfen, ob die Datei Excel oder CSV ist
        if uploaded_file.name.endswith('.xlsx') or uploaded_file.name.endswith('.xls'):
            # Excel-Datei laden und Blatt 'Touren' lesen
            df = pd.read_excel(uploaded_file, sheet_name="Touren")
            st.success("Das Blatt 'Touren' wurde erfolgreich geladen!")
        else:
            # CSV-Datei laden
            df = pd.read_csv(uploaded_file)
            st.success("CSV-Datei wurde erfolgreich geladen!")

        # Anzeige der ursprünglichen Daten
        st.write("Originaldaten:")
        st.dataframe(df.style.set_properties(**{
            'background-color': '#f4f4f4',
            'border': '1px solid #ddd',
            'color': '#333',
            'font-size': '12px',
            'text-align': 'center'
        }))

        # **Automatische Suchoptionen**
        search_numbers = ["602", "620", "350", "520", "156"]  # Zahlen, nach denen in 'Unnamed: 11' gesucht wird
        search_strings = ["AZ"]  # Nur nach "AZ" in 'Unnamed: 14' suchen

        # Prüfen, ob die Spalten vorhanden sind
        required_columns = ['Unnamed: 0', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 6',
                            'Unnamed: 7', 'Unnamed: 11', 'Unnamed: 12', 'Unnamed: 14']

        if all(col in df.columns for col in required_columns):
            # Suche nach den Zahlen in 'Unnamed: 11', aber schließe 607 aus
            number_matches = df[df['Unnamed: 11'].astype(str).isin(search_numbers)]

            # Suche nach "AZ" in 'Unnamed: 14'
            text_matches = df[df['Unnamed: 14'].str.contains('|'.join(search_strings), case=False, na=False)]

            # Kombinieren der Suchergebnisse
            combined_results = pd.concat([number_matches, text_matches]).drop_duplicates()

            # 607 aus allen Ergebnissen ausschließen
            combined_results = combined_results[combined_results['Unnamed: 11'].astype(str) != "607"]

            # Nur die gewünschten Spalten extrahieren und umbenennen
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

            # Wertzuweisung basierend auf Kennzeichen
            def calculate_earnings(kennzeichen):
                if kennzeichen in ["602", "156"]:
                    return 40
                elif kennzeichen in ["620", "350", "520"]:
                    return 20
                return 0

            final_results['Verdienst (€)'] = final_results['Kennzeichen'].astype(str).apply(calculate_earnings)

            # NaN-Werte durch leere Strings oder Nullen ersetzen
            final_results = final_results.fillna('')

            # Zusammenfassung des Verdienstes pro Fahrer
            earnings_summary = final_results.groupby(['Nachname', 'Vorname'], as_index=False)['Verdienst (€)'].sum()
            earnings_summary = earnings_summary.rename(columns={'Verdienst (€)': 'Gesamtverdienst (€)'})

            # **Sortieren nach Nachname und Vorname**
            final_results = final_results.sort_values(by=['Nachname', 'Vorname'])
            earnings_summary = earnings_summary.sort_values(by=['Nachname', 'Vorname'])

            # Export in Excel-Datei
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                workbook = writer.book
                worksheet1 = workbook.add_worksheet("Suchergebnisse")
                worksheet2 = workbook.add_worksheet("Zusammenfassung")

                writer.sheets["Suchergebnisse"] = worksheet1
                writer.sheets["Zusammenfassung"] = worksheet2

                worksheet1.write(0, 0, f"Kalenderwoche: {kalenderwoche}")

                # Schreibe die Daten ab der zweiten Zeile
                final_results.to_excel(writer, index=False, sheet_name="Suchergebnisse", startrow=2)
                earnings_summary.to_excel(writer, index=False, sheet_name="Zusammenfassung", startrow=0)

                # Lesbarkeit verbessern und Farben hinzufügen
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'top',
                    'fg_color': '#D9EAD3',
                    'border': 1
                })

                cell_format = workbook.add_format({
                    'border': 1,
                    'text_wrap': True,
                    'valign': 'top',
                    'fg_color': '#FFFAF0'
                })

                for col_num, value in enumerate(final_results.columns):
                    worksheet1.write(2, col_num, value, header_format)
                    worksheet1.set_column(col_num, col_num, 20)
                for row_num, row_data in final_results.iterrows():
                    for col_num, value in enumerate(row_data):
                        worksheet1.write(row_num + 3, col_num, value if pd.notnull(value) else '', cell_format)

                for col_num, value in enumerate(earnings_summary.columns):
                    worksheet2.write(0, col_num, value, header_format)
                    worksheet2.set_column(col_num, col_num, 20)
                for row_num, row_data in earnings_summary.iterrows():
                    for col_num, value in enumerate(row_data):
                        worksheet2.write(row_num + 1, col_num, value if pd.notnull(value) else '', cell_format)

            st.download_button(
                label="Suchergebnisse und Zusammenfassung als Excel herunterladen",
                data=output.getvalue(),
                file_name="Suchergebnisse_mit_Zusammenfassung.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    except Exception as e:
        st.error(f"Fehler beim Verarbeiten der Datei: {e}")
else:
    st.info("Bitte lade eine Excel- oder CSV-Datei hoch, um zu starten.")
