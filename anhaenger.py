import streamlit as st
import pandas as pd

# Titel der App
st.title("Touren-Such-App")
st.write("Lade eine Datei hoch, und die Daten werden in den Spalten `Unnamed: 11` und `Unnamed: 14` automatisch durchsucht.")

# Datei-Upload
uploaded_file = st.file_uploader("Lade deine Excel- oder CSV-Datei hoch", type=["xlsx", "xls", "csv"])

if uploaded_file:
    try:
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

        # Prüfen, ob die Spalten 'Unnamed: 11' und 'Unnamed: 14' vorhanden sind
        if 'Unnamed: 11' in df.columns and 'Unnamed: 14' in df.columns:
            # Suche nach den Zahlen in 'Unnamed: 11'
            number_matches = df[df['Unnamed: 11'].astype(str).isin(search_numbers)]

            # Suche nach den Zeichenfolgen in 'Unnamed: 14'
            text_matches = df[df['Unnamed: 14'].str.contains('|'.join(search_strings), case=False, na=False)]

            # Kombinieren der Suchergebnisse
            combined_results = pd.concat([number_matches, text_matches]).drop_duplicates()

            # Suchergebnisse anzeigen
            st.write("Suchergebnisse:")
            if not combined_results.empty:
                st.dataframe(combined_results)

                # Export der Suchergebnisse
                st.download_button(
                    label="Suchergebnisse als CSV herunterladen",
                    data=combined_results.to_csv(index=False).encode('utf-8'),
                    file_name="suchergebnisse.csv",
                    mime="text/csv",
                )
            else:
                st.warning("Keine Treffer gefunden.")
        else:
            st.error("Die benötigten Spalten 'Unnamed: 11' und 'Unnamed: 14' fehlen in der Datei.")

    except Exception as e:
        st.error(f"Fehler beim Verarbeiten der Datei: {e}")
else:
    st.info("Bitte lade eine Excel- oder CSV-Datei hoch, um zu starten.")
