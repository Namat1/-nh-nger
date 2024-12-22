import streamlit as st
import pandas as pd
from io import BytesIO

def filter_tours(file):
    # Excel-Datei einlesen
    df = pd.read_excel(file, sheet_name="Touren", engine="openpyxl", header=0)

    # Spaltennamen bereinigen
    df.columns = df.columns.str.strip()  # Entfernt führende/nachfolgende Leerzeichen
    df.columns = df.columns.str.replace(r"\s+", " ", regex=True)  # Ersetzt mehrfachen Leerraum
    df.columns = df.columns.str.lower()  # Macht alle Spaltennamen kleinbuchstabig

    # Erwartete Spaltennamen (nach Bereinigung)
    required_columns = ['l', 'o', 'a', 'd', 'e', 'g', 'h']

    # Prüfen, ob die bereinigten Spalten vorhanden sind
    missing_columns = [col for col in required_columns if col not in df.columns]

    if missing_columns:
        st.error(f"Die folgenden Spalten fehlen in der Datei: {', '.join(missing_columns)}")
        st.stop()

    # Filterkriterien definieren
    numbers_to_search = ["602", "156", "620", "350", "520"]
    az_mw_values = ["az", "mw"]

    # Zeilen filtern
    filtered_df = df[(df['l'].isin(numbers_to_search)) & (df['o'].str.lower().isin(az_mw_values))]

    # Werte aus den relevanten Spalten holen
    result = []
    for _, row in filtered_df.iterrows():
        tour = row['a']
        if pd.notna(row['d']) and pd.notna(row['e']):
            name = f"{row['d']} {row['e']}"
        elif pd.notna(row['g']) and pd.notna(row['h']):
            name = f"{row['g']} {row['h']}"
        else:
            name = "Unbekannt"

        result.append([tour, name])

    # Neue Tabelle erstellen
    result_df = pd.DataFrame(result, columns=["Tournummer", "Name"])
    return result_df

def convert_df_to_excel(df):
    # DataFrame in eine Excel-Datei umwandeln
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Gefilterte Touren')
    processed_data = output.getvalue()
    return processed_data

# Streamlit App
st.title("Touren Filter und Export")

uploaded_file = st.file_uploader("Laden Sie eine Excel-Datei hoch", type="xlsx")

if uploaded_file:
    # Filterprozess starten
    filtered_data = filter_tours(uploaded_file)

    # Gefilterte Daten anzeigen
    st.dataframe(filtered_data)

    # Möglichkeit zum Download der Ergebnisse
    excel_data = convert_df_to_excel(filtered_data)
    st.download_button(
        label="Download Excel Datei",
        data=excel_data,
        file_name="Gefilterte_Touren.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
