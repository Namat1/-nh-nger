import streamlit as st
import pandas as pd
from io import BytesIO

def filter_tours(file):
    # Excel-Datei einlesen ohne Header
    df = pd.read_excel(file, sheet_name="Touren", engine="openpyxl", header=None)  # Header wird ignoriert

    # Feste Buchstaben als Spaltennamen zuweisen
    max_columns = df.shape[1]  # Anzahl der Spalten bestimmen
    column_names = [chr(65 + i) for i in range(max_columns)]  # Generiere Spaltennamen A, B, C, ...
    df.columns = column_names  # Weist die Spaltennamen zu

    # Spaltenüberschriften anzeigen
    st.write("Zugewiesene Spaltennamen:", column_names)

    # Filterkriterien definieren (basierend auf Buchstaben)
    numbers_to_search = ["602", "156", "620", "350", "520"]
    az_mw_values = ["AZ", "Az", "az", "MW", "Mw", "mw"]

    # Zeilen filtern
    try:
        filtered_df = df[(df['L'].isin(numbers_to_search)) & (df['O'].str.strip().isin(az_mw_values))]
    except KeyError as e:
        st.error(f"Die Spalte {e} existiert nicht. Überprüfen Sie, ob die Datei korrekt formatiert ist.")
        st.stop()

    # Werte aus den relevanten Spalten holen
    result = []
    for _, row in filtered_df.iterrows():
        tour = row['A']  # Tournummer in Spalte A
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
    st.write("Datei erfolgreich hochgeladen. Verarbeite Daten...")
    filtered_data = filter_tours(uploaded_file)

    # Gefilterte Daten anzeigen
    st.write("Gefilterte Touren:")
    st.dataframe(filtered_data)

    # Möglichkeit zum Download der Ergebnisse
    excel_data = convert_df_to_excel(filtered_data)
    st.download_button(
        label="Download Excel Datei",
        data=excel_data,
        file_name="Gefilterte_Touren.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
