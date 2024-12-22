import streamlit as st
import pandas as pd
import re

# Titel der App
st.title("Touren-Filter und Such-App")
st.write("Lade eine Excel-Datei hoch, exportiere das Blatt 'Touren', filtere nach Kriterien und suche nach Zahlen oder Textmustern.")

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

        # **Filterung der Daten**
        st.sidebar.header("Filteroptionen")
        filtered_df = df.copy()

        for column in df.columns:
            if df[column].dtype == 'object':  # Textspalten
                filter_value = st.sidebar.text_input(f"Filter für {column}", "")
                if filter_value:
                    filtered_df = filtered_df[filtered_df[column].str.contains(filter_value, case=False, na=False)]

            elif pd.api.types.is_numeric_dtype(df[column]):  # Numerische Spalten
                min_val = st.sidebar.number_input(f"Min {column}", value=float(df[column].min()), step=1.0)
                max_val = st.sidebar.number_input(f"Max {column}", value=float(df[column].max()), step=1.0)
                filtered_df = filtered_df[(filtered_df[column] >= min_val) & (filtered_df[column] <= max_val)]

        # Gefilterte Daten anzeigen
        st.write("Gefilterte Daten:")
        st.dataframe(filtered_df)

        # **Suche nach Zahlen und Textmustern**
        st.sidebar.header("Suchoptionen")
        search_numbers = st.sidebar.text_input("Gib die Zahlen (kommagetrennt) ein, nach denen gesucht werden soll:", "602,620,350,520,156")
        search_strings = st.sidebar.text_input("Gib die Zeichenfolgen (kommagetrennt) ein, nach denen gesucht werden soll:", "AZ,Az,az,MW,Mw,mw")

        # Umwandlung der Eingaben in Listen
        number_patterns = [num.strip() for num in search_numbers.split(",")]
        string_patterns = [string.strip() for string in search_strings.split(",")]

        # Suche in den gefilterten Daten
        results = []
        for index, row in filtered_df.iterrows():
            row_content = " ".join(row.astype(str).values)  # Zeileninhalt als String
            if any(re.search(rf"\b{num}\b", row_content) for num in number_patterns) or \
               any(re.search(rf"{string}", row_content, re.IGNORECASE) for string in string_patterns):
                results.append(row)

        # Ergebnisse in einen DataFrame umwandeln
        search_results = pd.DataFrame(results)

        # Anzeige der Suchergebnisse
        st.write("Suchergebnisse:")
        if not search_results.empty:
            st.dataframe(search_results)
            
            # Export der Suchergebnisse
            st.download_button(
                label="Suchergebnisse als CSV herunterladen",
                data=search_results.to_csv(index=False).encode('utf-8'),
                file_name="suchergebnisse.csv",
                mime="text/csv",
            )
        else:
            st.warning("Keine Treffer gefunden.")

    except Exception as e:
        st.error(f"Fehler beim Verarbeiten der Datei: {e}")
else:
    st.info("Bitte lade eine Excel- oder CSV-Datei hoch, um zu starten.")
