import streamlit as st
import pandas as pd
from io import BytesIO

def analyze_and_display(file):
    # Excel-Datei einlesen (Header ab Zeile 5, da die relevanten Daten dort beginnen)
    df = pd.read_excel(file, sheet_name="Touren", engine="openpyxl", header=4)

    # Spalten anzeigen
    st.write("Gefundene Spalten:")
    st.write(df.columns.tolist())

    # Erstelle eine Tabelle mit Spaltenpositionen und Beispielwerten
    column_analysis = pd.DataFrame({
        "Spaltenname": df.columns,
        "Spaltenposition (Excel)": [chr(65 + i) for i in range(len(df.columns))],
        "Erster Wert (Zeile 6)": df.iloc[0].tolist()
    })

    # Analysedaten anzeigen
    st.write("Analyse der Spalten:")
    st.dataframe(column_analysis)

    # Zeige die ersten Zeilen des DataFrames
    st.write("Erste 10 Zeilen der Daten:")
    st.dataframe(df.head(10))

    return df

# Streamlit App
st.title("Analyse der Excel-Daten")

uploaded_file = st.file_uploader("Laden Sie eine Excel-Datei hoch", type="xlsx")

if uploaded_file:
    st.write("Datei erfolgreich hochgeladen. Analysiere Daten...")
    df = analyze_and_display(uploaded_file)
