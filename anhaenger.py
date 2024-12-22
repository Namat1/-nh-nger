import streamlit as st
import pandas as pd
from io import BytesIO

def analyze_and_prepare_download(file):
    try:
        # Excel-Datei einlesen (Header ab Zeile 5)
        df = pd.read_excel(file, sheet_name="Touren", engine="openpyxl", header=4)

        # Spalten bereinigen
        df.columns = df.columns.str.strip()
        df.columns = df.columns.str.replace(r"\s+", " ", regex=True)

        # Entferne komplett leere Spalten und Zeilen
        df.dropna(how='all', axis=0, inplace=True)  # Leere Zeilen entfernen
        df.dropna(how='all', axis=1, inplace=True)  # Leere Spalten entfernen

        # Zeige Spaltenanalyse
        column_analysis = pd.DataFrame({
            "Spaltenname": df.columns,
            "Spaltenposition (Excel)": [chr(65 + i) for i in range(len(df.columns))],
            "Erster Wert (Zeile 6)": df.iloc[0].tolist() if not df.empty else []
        })

        st.write("Analyse der Spalten:")
        st.dataframe(column_analysis)

        st.write("Erste 10 Zeilen der Daten:")
        st.dataframe(df.head(10))

        # Erstelle eine Excel-Datei für den Download
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            column_analysis.to_excel(writer, sheet_name="Spaltenanalyse", index=False)
            df.to_excel(writer, sheet_name="Daten", index=False)
        processed_data = output.getvalue()

        return processed_data
    except ValueError as e:
        st.error(f"Fehler beim Verarbeiten der Daten: {e}")
        return None

# Streamlit App
st.title("Analyse der Excel-Daten mit Download")

uploaded_file = st.file_uploader("Laden Sie eine Excel-Datei hoch", type="xlsx")

if uploaded_file:
    st.write("Datei erfolgreich hochgeladen. Analysiere Daten...")
    excel_data = analyze_and_prepare_download(uploaded_file)

    if excel_data:
        st.download_button(
            label="Download analysierte Excel-Datei",
            data=excel_data,
            file_name="Analyisierte_Daten.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Die Datei konnte nicht verarbeitet werden. Überprüfen Sie die Datenstruktur.")
