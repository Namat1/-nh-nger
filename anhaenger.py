from openpyxl import load_workbook

# Excel-Datei laden
workbook = load_workbook("deine_datei.xlsx")
if "Touren" in workbook.sheetnames:
    sheet = workbook["Touren"]  # Blatt 'Touren' auswählen

    # Alle Zellen durchgehen und Farbwerte ermitteln
    print("Farbwerte der Zellen im Blatt 'Touren':")
    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
        for cell in row:
            if cell.fill and cell.fill.start_color:
                # Farbe der Zelle auslesen
                try:
                    color = cell.fill.start_color.rgb  # RGB-Wert der Zelle
                    print(f"Zelle {cell.coordinate} hat die Farbe: {color}")
                except AttributeError:
                    # Falls keine RGB-Farbe vorhanden ist
                    print(f"Zelle {cell.coordinate} hat keine Farbe.")
else:
    print("Das Blatt 'Touren' wurde in der Datei nicht gefunden.")
