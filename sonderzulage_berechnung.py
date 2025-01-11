def apply_styles(sheet):
    """
    Dynamische Anwendung eines klaren und übersichtlichen Business-Stils:
    - Namenszeilen: Hellblau, fett.
    - Kopfzeilen (Überschriften): Hellgrau, fett.
    - Gesamtverdienstzeilen: Hellgrün, fett.
    - Datenzeilen: Weiß, normal.
    - Werte in der Spalte "Verdienst" behalten das €-Zeichen als benutzerdefiniertes Format.
    """
    # Stildefinitionen
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    name_fill = PatternFill(start_color="D9EAF7", end_color="D9EAF7", fill_type="solid")  # Hellblau für Namenszeilen
    header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")  # Hellgrau für Kopfzeilen
    total_fill = PatternFill(start_color="DFF7DF", end_color="DFF7DF", fill_type="solid")  # Hellgrün für Gesamtverdienstzeilen
    data_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")  # Weiß für Datenzeilen

    for row in sheet.iter_rows():
        first_cell_value = str(row[0].value).strip() if row[0].value else ""

        if "Gesamtverdienst" in first_cell_value:  # Gesamtverdienstzeilen
            for cell in row:
                cell.fill = total_fill
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center")
                cell.border = thin_border
                if cell.column == 5 and isinstance(cell.value, (int, float)):  # Spalte "Verdienst" (5. Spalte)
                    cell.number_format = '#,##0.00 €'

        elif first_cell_value and any(char.isalpha() for char in first_cell_value) and not "Datum" in first_cell_value:  # Namenszeilen
            for cell in row:
                cell.fill = name_fill
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center")
                cell.border = thin_border

        elif "Datum" in first_cell_value:  # Kopfzeilen (Überschriften)
            for cell in row:
                cell.fill = header_fill
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center")
                cell.border = thin_border

        else:  # Datenzeilen
            for cell in row:
                cell.fill = data_fill
                cell.font = Font(bold=False)
                cell.alignment = Alignment(horizontal="left")
                cell.border = thin_border
                if cell.column == 5 and isinstance(cell.value, (int, float)):  # Spalte "Verdienst" (5. Spalte)
                    cell.number_format = '#,##0.00 €'
