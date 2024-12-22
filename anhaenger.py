import streamlit as st
import pandas as pd
from io import BytesIO

def filter_tours(file):
    # Excel-Datei einlesen (Header ab Zeile 5)
    df = pd.read_excel(file, sheet_name="Touren", engine="openpyxl", header=4)

    # Spalten bereinigen
    df.columns = df.columns.str.strip()
    df.columns = df.columns.str.replace(r"\s+", " ", regex=True)

    # Spalten korrekt zuordnen
    spalte_a = "Unnamed: 0"  # Spalte A (Tournummer)
    spalte_d = "Pekrul"      # Spalte D (Name Teil 1)
    spalte_e = "Olaf"        # Spalte E (Name Teil 2)
    spalte_g = "Unnamed: 6"  # Spalte G (Name Teil 3)
    spalte_h = "Unnamed: 7"  # Spalte H (Name Teil 4)
    spalte_l = "Unnamed: 11" # Spalte L (Filterkriterium)
    spalte_o = "Unnamed: 14" # Spalte O (Filterkriterium)

    # Debug: Inhalt der relevanten Spalten anzeigen
    st.write("Inhalt der Spalte L (erste 20 Werte):", df[spalte_l].head(20).tolist())
    st.write("Inhalt der Spalte O (erste 20 Werte):", df[spalte_o].head(20).tolist())

    # Bereinige die relevanten Spalten
    df[spalte_l] = df[spalte_l].astype(str).str.strip()
    df[spalte_o] = df[spalte_o].astype(str).str.strip().str.upper()

    # Einzelne Filter prüfen
    filtered_by_l = df[df[spalte_l].isin(["602", "156", "620", "350", "520"])]
    
    if filtered_by_l.empty:
        st.error("Der DataFrame nach Filterung auf Spalte L ist leer.")
    else:
        st.write(f"Nach Filter auf Spalte L: {len(filtered_by_l)} Zeilen gefunden")
        st.dataframe(filtered_by_l.head(10))

    filtered_by_o = df[df[spalte_o].isin(["AZ", "MW"])]
    if filtered_by_o.empty:
        st.error("Der DataFrame nach Filterung auf Spalte O ist leer.")
    else:
        st.write(f"Nach Filter auf Spalte O: {len(filtered_by_o)} Zeilen gefunden")
        st.dataframe(filtered_by_o.head(10))

    # Gesamte Filterung
    filtered_df = df[
        (df[spalte_l].isin(["602", "156", "620", "350", "520"])) & 
        (df[spalte_o].isin(["AZ", "MW"]))
    ]

    # Prüfen, ob mindestens 500 Zeilen im Ergebnis sind
    if len(filtered_df) < 500:
        st.error(f"Die Filterung ergab nur {len(filtered_df)} Zeilen. Es werden mindestens 500 Zeilen benötigt.")
        return None

    # Werte aus den relevanten Spalten holen
    result = []
    for _, row in filtered_df.iterrows():
        tour = row[spalte_a]
        if pd.notna(row[spalte_d]) and pd.notna(row[spalte_e]):
            name = f"{row[spalte_d]} {row[spalte_e]}"
        elif pd.notna(row[spalte_g]) and pd.notna(row[spalte_h]):
            name = f"{row[spalte_g]} {row[spalte_h]}"
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
st.title("Touren Filter und Export (mind. 500 Zeilen)")

uploaded_file = st.file_uploader("Laden Sie eine Excel-Datei hoch", type="xlsx")

if uploaded_file:
    st.write("Datei erfolgreich hochgeladen. Verarbeite Daten...")
    filtered_data = filter_tours(uploaded_file)

    if filtered_data is not None:
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
    else:
        st.error("Die Filterung ergab nicht genügend Zeilen. Bitte überprüfen Sie die Daten.")
