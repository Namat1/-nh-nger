import streamlit as st
import pandas as pd
from io import BytesIO

def filter_tours(file):
    # Excel-Datei einlesen
    try:
        df = pd.read_excel(file, sheet_name="Touren", engine="openpyxl")
    except ValueError as e:
        st.error("Tabellenblatt 'Touren' konnte nicht gefunden werden. Überprüfen Sie die Datei.")
        st.stop()

    # Zeige die ersten Zeilen und Spaltennamen für Debugging
    st.write("Erste Zeilen der Datei:")
    st.dataframe(df.head())
    st.write("Gefundene Spalten:")
    st.write(df.columns.tolist())

    # Spaltennamen bereinigen
    df.columns = df.columns.str.strip()

    # Sicherstellen, dass alle Werte Strings sind
    df = df.astype(str)

    # Prüfen, ob die benötigten Spalten vorhanden sind
    required_columns = ['L', 'O', 'A', 'D', 'E', 'G', 'H']
    missing_columns = [col for col in required_columns if col not in df.columns]

    if missing_columns:
        st.error(f"Die folgenden Spalten fehlen in der Datei: {', '.join(missing_columns)}")
        st.stop()

    # Filterkriterien definieren
    numbers_to_search = ["602", "156", "620", "350", "520"]
    az_mw_values = ["AZ", "Az", "az", "MW", "Mw", "mw"]

    # Debugging: Prüfe die Inhalte der relevanten Spalten
    st.write("Beispielwerte aus Spalte 'L':")
    st.write(df['L'].unique())
    st.write("Beispielwerte aus Spalte 'O':")
    st.write(df['O'].unique())

    # Zeilen filtern
    try:
        filtered_df = df[(df['L'].isin(numbers_to_search)) & (df['O'].isin(az_mw_values))]
    except Exception as e:
        st.error(f"Fehler beim Filtern der Zeilen: {e}")
        st.stop()

    # Debugging: Zeige gefilterte Daten
    st.write("Gefilterte Daten:")
    st.dataframe(filtered_df)

    # Werte aus den relevanten Spalten holen
    result = []
    for _, row in filtered_df.iterrows():
        tour = row['A']
        if pd.notna(row['D']) and pd.notna(row['E']):
            name = f"{row['D']} {row['E']}"
        elif pd.notna(row['G']) and pd.notna(row['H']):
            name = f"{row['G']} {row['H']}"
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
    st.write("Datei erfolgreich hochgeladen. Verarbeite Daten...")
    filtered_data = filter_tours(uploaded_file)

    # Gefilterte Daten anzeigen
    st.write("Gefilterte Touren:")
    st.dataframe(filtered_data)

    # Möglichkeit zum Download der Ergebnisse
    st.write("Laden Sie die gefilterten Daten als Excel-Datei herunter:")
    excel_data = convert_df_to_excel(filtered_data)
    st.download_button(
        label="Download Excel Datei",
        data=excel_data,
        file_name="Gefilterte_Touren.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
