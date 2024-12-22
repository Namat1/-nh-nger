import streamlit as st
from openpyxl import load_workbook

# Titel der App
st.title("Zellenfarben im Blatt 'Touren' auswerten")

# Datei-Upload
uploaded_file = st.file_uploader("Lade deine Excel-Datei hoch", type=["xlsx", "xls"])

if uploaded_file:
    try:
        # Excel-Datei laden
        workbook = load_workbook(uploaded_file, data_only=True)

        # Prüfen, ob das Blatt 'Touren' existiert
        if "Touren" in workbook.sheetnames:
            sheet = workbook["Touren"]  # Blatt 'Touren' auswählen

            # Ergebnis-Container
            color_results = []

            # Alle Zellen durchgehen und Farbwerte ermitteln
            for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
                for cell in row:
                    if cell.fill and cell.fill.start_color:
                        try:
                            # Farbe der Zelle auslesen
                            color = cell.fill.start_color.rgb  # RGB-Wert der Zelle
                            color_results.append((cell.coordinate, color))
                        except AttributeError:
                            color_results.append((cell.coordinate, "Keine Farbe"))

            # Ergebnisse anzeigen
            if color_results:
                st.write("Farbwerte der Zellen im Blatt 'Touren':")
                for coordinate, color in color_results:
                    st.write(f"Zelle {coordinate}: {color}")
            else:
                st.warning("Keine farbigen Zellen im Blatt 'Touren' gefunden.")
        else:
            st.error("Das Blatt 'Touren' wurde in der Datei nicht gefunden.")
    except Exception as e:
        st.error(f"Fehler beim Verarbeiten der Datei: {e}")
else:
    st.info("Bitte lade eine Excel-Datei hoch, um die Farbwerte zu ermitteln.")
