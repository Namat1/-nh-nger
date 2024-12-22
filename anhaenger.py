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
        st.dataframe(df)

        # **Automatische Suchoptionen**
        search_numbers = ["602", "620", "350", "520", "156"]  # Zahlen, nach denen in 'Unnamed: 11' gesucht wird
        search_strings = ["AZ", "Az", "az", "MW", "Mw", "mw"]  # Zeichenfolgen, nach denen in 'Unnamed: 14' gesucht wird

        # Prüfen, ob die Spalten vorhanden sind
        required_columns = ['Unnamed: 0', 'Unnamed: 3', 'Unnamed: 4', 'Unnamed: 6',
                            'Unnamed: 7', 'Unnamed: 11', 'Unnamed: 12', 'Unnamed: 14']

        if all(col in df.columns for col in required_columns):
            # Suche nach den Zahlen in 'Unnamed: 11', wobei 607 vollständig ausgeschlossen wird
            number_matches = df[
                df['Unnamed: 11'].astype(str).isin(search_numbers) & 
                (df['Unnamed: 11'].astype(str) != "607")
            ]

            # Suche nach den Zeichenfolgen in 'Unnamed: 14', wobei Zeilen mit 607 ausgeschlossen werden
            text_matches = df[
                df['Unnamed: 14'].str.contains('|'.join(search_strings), case=False, na=False) &
                (df['Unnamed: 11'].astype(str) != "607")
            ]

            # Kombinieren der Suchergebnisse
            combined_results = pd.concat([number_matches, text_matches]).drop_duplicates()

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

            # Sortieren nach Nachname und Vorname
            final_results = final_results.sort_values(by=['Nachname', 'Vorname'])

            # Verdienstberechnung
            payment_mapping = {
                "602": 40,
                "156": 40,
                "620": 20,
                "350": 20,
                "520": 20
            }
            final_results['Verdienst'] = final_results['Kennzeichen'].map(payment_mapping).fillna(0)

            # Tabellarische Zusammenfassung
            summary = final_results.groupby(['Nachname', 'Vorname'])['Verdienst'].sum().reset_index()
            summary = summary.rename(columns={"Verdienst": "Gesamtverdienst"})

            # Suchergebnisse anzeigen
            st.write("Suchergebnisse:")
            if not final_results.empty:
                st.dataframe(final_results)

                # Export in Excel-Datei
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    # Schreibe Kalenderwoche in die erste Zeile
                    workbook = writer.book

                    # Blatt mit Suchergebnissen
                    worksheet = workbook.add_worksheet("Suchergebnisse")
                    writer.sheets["Suchergebnisse"] = worksheet
                    worksheet.write(0, 0, f"Kalenderwoche: {kalenderwoche}")
                    final_results.to_excel(writer, index=False, sheet_name="Suchergebnisse", startrow=2)
                    for i, column in enumerate(final_results.columns):
                        column_width = max(final_results[column].astype(str).map(len).max(), len(column))
                        worksheet.set_column(i, i, column_width)

                    # Blatt mit Zusammenfassung
                    summary_worksheet = workbook.add_worksheet("Zusammenfassung")
                    writer.sheets["Zusammenfassung"] = summary_worksheet
                    summary.to_excel(writer, index=False, sheet_name="Zusammenfassung", startrow=0)
                    for i, column in enumerate(summary.columns):
                        column_width = max(summary[column].astype(str).map(len).max(), len(column))
                        summary_worksheet.set_column(i, i, column_width)

                # Export-Button für Excel-Datei
                st.download_button(
                    label="Suchergebnisse und Zusammenfassung als Excel herunterladen",
                    data=output.getvalue(),
                    file_name="Suchergebnisse_mit_Zusammenfassung.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.warning("Keine Treffer gefunden.")
        else:
            missing_columns = [col for col in required_columns if col not in df.columns]
            st.error(f"Die folgenden benötigten Spalten fehlen in der Datei: {', '.join(missing_columns)}")

    except Exception as e:
        st.error(f"Fehler beim Verarbeiten der Datei: {e}")
else:
    st.info("Bitte lade eine Excel- oder CSV-Datei hoch, um zu starten.")
