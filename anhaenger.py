import streamlit as st
import pandas as pd
from io import BytesIO

def analyze_row_13(file):
    try:
        # Excel-Datei einlesen (Header ab Zeile 5)
        df = pd.read_excel(file, sheet_name="Touren", engine="openpyxl", header=4)

        # Spalten bereinigen
        df.columns = df.columns.str.strip()
        df.columns = df.columns.str.replace(r"\s+", " ", regex=True)

        # Entferne komplett leere Spalten und Zeilen
        df.dropna(how='all', axis=0, inplace=True)
        df.dropna(how='all', axis=1, inplace=True)

        # Spalten korrekt zuordnen
        spalte_a = "Unnamed: 0"  # Spalte A (Tournummer)
        spalte_d = "Pekrul"      # Spalte D (Name Teil 1)
        spalte_e = "Olaf"        # Spalte E (Name Teil 2)
        spalte_g = "Unnamed: 6"  # Spalte G (Name Teil 3)
        spalte_h = "Unnamed: 7"  # Spalte H (Name Teil 4)
        spalte_l = "Unnamed: 11" # Spalte L (Filterkriterium)
        spalte_o = "Unnamed: 14" # Spalte O (Filterkriterium)

        # Prüfen, ob Zeile 13 existiert
        if len(df) >= 9:  # Zeile 13 in Excel ist Zeile 9 im DataFrame (0-basierter Index)
            row_13 = df.iloc[8]  # Zeile 13 auslesen (Index 8)
            st.write("Auswertung von Zeile 13:")
            st.write(row_13)

            # Erstelle eine Tabelle mit Spaltennamen und Werten aus Zeile 13
            row_13_analysis = pd.DataFrame({
                "Spaltenname": df.columns,
                "Wert in Zeile 13": row_13.tolist()
            })
            st.write("Analyse von Zeile 13:")
            st.dataframe(row_13_analysis)

            # Werte aus den relevanten Spalten von Zeile 13 extrahieren
            tournummer = row_13[spalte_a]
            name = (
                f"{row_13[spalte_d]} {row_13[spalte_e]}"
                if pd.notna(row_13[spalte_d]) and pd.notna(row_13[spalte_e])
                else f"{row_13[spalte_g]} {row_13[spalte_h]}"
                if pd.notna(row_13[spalte_g]) and pd.notna(row_13[spalte_h])
                else "Unbekannt"
            )
            filter_l = row_13[spalte_l]
            filter_o = row_13[spalte_o]

            st.write(f"Tournummer: {tournummer}")
            st.write(f"Name: {name}")
            st.write(f"Wert in Spalte L: {filter_l}")
            st.write(f"Wert in Spalte O: {filter_o}")
        else:
            st.error("Zeile 13 existiert nicht in den eingelesenen Daten.")

        # Optional: Datei für den Download vorbereiten
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            row_13_analysis.to_excel(writer, sheet_name="Zeile_13", index=False)
            df.to_excel(writer, sheet_name="Daten", index=False)
        processed_data = output.getvalue()

        return processed_data
    except Exception as e:
        st.error(f"Fehler bei der Verarbeitung: {e}")
        return None

# Streamlit App
st.title("Analyse von Zeile 13 in Excel")

uploaded_file = st.file_uploader("Laden Sie eine Excel-Datei hoch", type="xlsx")

if uploaded_file:
    st.write("Datei erfolgreich hochgeladen. Analysiere Zeile 13...")
    excel_data = analyze_row_13(uploaded_file)

    if excel_data:
        st.download_button(
            label="Download analysierte Excel-Datei",
            data=excel_data,
            file_name="Zeile_13_Analyse.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Die Datei konnte nicht verarbeitet werden.")
