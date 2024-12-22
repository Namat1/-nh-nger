import streamlit as st
import pandas as pd
import re

# Titel der App
st.title("Touren-Filter und Such-App")
st.write("Lade eine Datei hoch, und die Daten werden automatisch durchsucht.")

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
        search_numbers = ["602", "620", "350", "520", "156"]  # Zahlen, nach denen gesucht wird
        search_strings = ["AZ", "Az", "az", "MW", "Mw", "mw"]  # Zeichenfolgen, nach denen gesucht wird

        # Automatische Suche in den Daten
        search_results = []
        for _, row in df.iterrows():
            row_content = " ".join(row.astype(str).values)  # Zeileninhalt als ein String
            # Überprüfen auf Zahlen und Strings
            if any(re.search(rf"\b{num}\b", row_content) for num in search_numbers) or \
               any(re.search(rf"{string}", row_content, re.IGNORECASE) for string in search_strings):
                search_results.append(row)

        # Suchergebnisse anzeigen
        search_results_df = pd.DataFrame(search_results)
        st.write("Suchergebnisse:")
        if not search_results_df.empty:
            st.dataframe(search_results_df)
            
            # Export der Suchergebnisse
            st.download_button(
                label="Suchergebnisse als CSV herunterladen",
                data=search_results_df.to_csv(index=False).encode('utf-8'),
                file_name="suchergebnisse.csv",
                mime="text/csv",
            )
        else:
            st.warning("Keine Treffer gefunden.")

    except Exception as e:
        st.error(f"Fehler beim Verarbeiten der Datei: {e}")
else:
    st.info("Bitte lade eine Excel- oder CSV-Datei hoch, um zu starten.")
