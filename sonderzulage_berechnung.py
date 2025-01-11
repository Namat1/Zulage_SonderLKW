import pandas as pd
import streamlit as st
import calendar
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# Personalnummer-Zuordnung
name_to_personalnummer = {
    "Adler": {"Philipp": "00041450"},
    "Scheil": {"Eric-Rene": "00038579", "Rene": "00020851"},
    # Weitere Zuordnungen ...
}

def apply_styles(sheet):
    """
    Dynamische Anwendung eines klaren und übersichtlichen Business-Stils:
    - Namenszeilen: Hellblau, fett, verbunden, mit Personalnummer.
    - Kopfzeilen (Überschriften): Hellgrau, fett.
    - Gesamtverdienstzeilen: Hellgrün, fett.
    - Datenzeilen: Weiß, normal.
    """
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    name_fill = PatternFill(start_color="D9EAF7", end_color="D9EAF7", fill_type="solid")  # Hellblau für Namenszeilen
    header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")  # Hellgrau für Kopfzeilen
    total_fill = PatternFill(start_color="DFF7DF", end_color="DFF7DF", fill_type="solid")  # Hellgrün für Gesamtverdienstzeilen
    data_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")  # Weiß für Datenzeilen

    for row_idx, row in enumerate(sheet.iter_rows(), start=1):
        first_cell_value = str(row[0].value).strip() if row[0].value else ""

        if row_idx == 1 and "Datum" not in first_cell_value:  # Überspringe ungültige erste Zeile
            continue

        if "Gesamtverdienst" in first_cell_value:  # Gesamtverdienstzeilen
            for cell in row:
                cell.fill = total_fill
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="right")
                cell.border = thin_border
                if cell.column == 5 and isinstance(cell.value, (int, float)):  # Spalte "Verdienst" (5. Spalte)
                    cell.number_format = '#,##0.00 €'

        elif first_cell_value and any(char.isalpha() for char in first_cell_value) and not "Datum" in first_cell_value:  # Namenszeilen
            try:
                vorname, nachname = first_cell_value.split(" ", 1)
                vorname = "".join(vorname.strip().split()).title()
                nachname = "".join(nachname.strip().split()).title()

                # Versuche direkte Zuordnung und alternative Schreibweisen
                personalnummer = (
                    name_to_personalnummer.get(nachname, {}).get(vorname)
                    or name_to_personalnummer.get(nachname, {}).get(vorname.replace("-", " "))  # Ohne Bindestrich
                    or name_to_personalnummer.get(nachname, {}).get(vorname.replace(" ", "-"))  # Mit Bindestrich
                    or "Unbekannt"
                )

            except ValueError:
                personalnummer = "Unbekannt"

            # Verbinden der Namenszeile über alle Spalten
            sheet.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=5)
            row[0].value = f"{first_cell_value} - {personalnummer}"  # Füge Personalnummer hinzu
            row[0].fill = name_fill
            row[0].font = Font(bold=True)
            row[0].alignment = Alignment(horizontal="center")
            row[0].border = thin_border
            for cell in row[1:]:
                cell.value = None  # Leere Zellen hinter dem Namen

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
                cell.alignment = Alignment(horizontal="right")
                cell.border = thin_border
                if cell.column == 5 and isinstance(cell.value, (int, float)):  # Spalte "Verdienst" (5. Spalte)
                    cell.number_format = '#,##0.00 €'


def main():
    st.title("Touren-Auswertung mit klarer Trennung der Namenszeile")

    uploaded_file = st.file_uploader("Lade eine Excel-Datei hoch", type=["xlsx"])
    if uploaded_file:
        df = pd.read_excel(uploaded_file, sheet_name="Touren", header=0)
        st.write("Überprüfe hochgeladene Daten:")
        st.write(df.head())  # Debugging: Zeigt die ersten Zeilen der Datei an

        # Entferne komplett leere Zeilen
        df = df.dropna(how="all")
        sheet_data = []

        # Verarbeite die Daten
        for _, row in df.iterrows():
            if pd.isnull(row["Datum"]):  # Ignoriere Zeilen ohne Datum
                continue

            # Verarbeite nur relevante Zeilen
            sheet_data.append([
                row["Datum"].strftime("%d.%m.%Y"),
                row["Tour"],
                row["LKW"],
                row["Art"],
                row["Verdienst"]
            ])

        output_file = "touren_auswertung_korrekt.xlsx"
        try:
            with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                workbook = writer.book
                sheet = workbook.active

                # Füge Daten in das Sheet ein
                for data_row in sheet_data:
                    sheet.append(data_row)

                # Wende Formatierungen an
                apply_styles(sheet)

            with open(output_file, "rb") as file:
                st.download_button(
                    label="Download Auswertung",
                    data=file,
                    file_name="touren_auswertung_korrekt.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Fehler beim Exportieren der Datei: {e}")


if __name__ == "__main__":
    main()
