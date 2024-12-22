import streamlit as st
import pandas as pd
from io import BytesIO

def analyze_row_15(file):
    try:
        # Excel-Datei einlesen (Header ab Zeile 5)
        df = pd.read_excel(file, sheet_name="Touren", engine="openpyxl", header=4)

        # Spalten bereinigen
        df.columns = df.columns.str.strip()
        df.columns = df.columns.str.replace(r"\s+", " ", regex=True)

        # Entferne komplett leere Spalten und Zeilen
        df.dropna(how='all', axis=0, inplace=True)
        df.dropna(how='all', axis=1, inplace=True)

        # Prüfen, ob Zeile 15 existiert
        if len(df) >= 11:  # Zeile 15 in Excel ist Zeile 11 im DataFrame (0-basierter Index)
            row_15 = df.iloc[10]  # Zeile 15 auslesen (Index 10)
            st.write("Auswertung von Zeile 15:")
            st.write(row_15)

            # Erstelle eine Tabelle mit Spaltennamen und Werten aus Zeile 15
            row_15_analysis = pd.DataFrame({
                "Spaltenname": df.columns,
                "Wert in Zeile 15": row_15.tolist()
            })
            st.write("Analyse von Zeile 15:")
            st.dataframe(row_15_analysis)
        else:
            st.error("Zeile 15 existiert nicht in den eingelesenen Daten.")

        # Zeige die ersten 10 Zeilen der Datei
        st.write("Erste 10 Zeilen der Daten:")
        st.dataframe(df.head(10))

        # Optional: Datei für den Download vorbereiten
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            row_15_analysis.to_excel(writer, sheet_name="Zeile_15", index=False)
            df.to_excel(writer, sheet_name="Daten", index=False)
        processed_data = output.getvalue()

        return processed_data
    except Exception as e:
        st.error(f"Fehler bei der Verarbeitung: {e}")
        return None

# Streamlit App
st.title("Analyse von Zeile 15 in Excel")

uploaded_file = st.file_uploader("Laden Sie eine Excel-Datei hoch", type="xlsx")

if uploaded_file:
    st.write("Datei erfolgreich hochgeladen. Analysiere Zeile 15...")
    excel_data = analyze_row_15(uploaded_file)

    if excel_data:
        st.download_button(
            label="Download analysierte Excel-Datei",
            data=excel_data,
            file_name="Zeile_15_Analyse.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Die Datei konnte nicht verarbeitet werden.")
