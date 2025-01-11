import pandas as pd
import streamlit as st
import calendar
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# Personalnummer-Zuordnung
name_to_personalnummer = {
    "Adler": {"Philipp": "00041450"},
    "Auer": {"Frank": "00020795"},
    "Batkowski": {"Tilo": "00046601"},
    "Benabbes": {"Badr": "00048980"},
    "Biebow": {"Thomas": "00042004"},
    "Bläsing": {"Elmar": "00049093"},
    "Bursian": {"Ronny": "00025714"},
    "Buth": {"Sven": "00046673"},
    "Böhnke": {"Marcel": "00020833"},
    "Carstensen": {"Martin": "00042412"},
    "Chege": {"Moses Gichuru": "00046106"},
    "Dammasch": {"Bernd": "00019297"},
    "Demuth": {"Harry": "00020796"},
    "Doroszkiewicz": {"Bogumil": "00049132"},
    "Dürr": {"Holger": "00039164"},
    "Effenberger": {"Sven": "00030807"},
    "Engel": {"Raymond": "00033429"},
    "Fechner": {"Danny": "00043696", "Klaus": "00038278"},
    "Findeklee": {"Bernd": "00020804"},
    "Flint": {"Henryk": "00042414"},
    "Fuhlbrügge": {"Justin": "00046289"},
    "Gehrmann": {"Rayk": "00046702"},
    "Gheonea": {"Costel-Daniel": "00050877"},
    "Glanz": {"Björn": "00041914"},
    "Gnech": {"Torsten": "00018613"},
    "Greve": {"Nicole": "00040760"},
    "Guthmann": {"Fred": "00018328"},
    "Hagen": {"Andy": "00020271"},
    "Hartig": {"Sebastian": "00044120"},
    "Haus": {"David": "00046101"},
    "Heeser": {"Bernd": "00041916"},
    "Helm": {"Philipp": "00046685"},
    "Henkel": {"Bastian": "00048187"},
    "Holtz": {"Torsten": "00021159"},
    "Janikiewicz": {"Radoslaw": "00042159"},
    "Kelling": {"Jonas Ole": "00044140"},
    "Kleiber": {"Lutz": "00026255"},
    "Klemkow": {"Ralf": "00040634"},
    "Kollmann": {"Steffen": "00040988"},
    "König": {"Heiko": "00036341"},
    "Krazewski": {"Cezary": "00039463"},
    "Krieger": {"Christian": "00049092"},
    "Krull": {"Benjamin": "00044192"},
    "Lange": {"Michael": "00035407"},
    "Lewandowski": {"Kamil": "00041044"},
    "Likoonski": {"Vladimir": "00044766"},
    "Linke": {"Erich": "00048377"},
    "Lefkih": {"Houssni": "00052293"},
    "Ludolf": {"Michel": "00048814"},
    "Marouni": {"Ayyoub": "00048986"},
    "Mintel": {"Mario": "00046686"},
    "Ohlenroth": {"Nadja": "00042114"},
    "Ohms": {"Torsten": "00019300"},
    "Okoth": {"Tedy Omondi": "00046107"},
    "Oszmian": {"Jacub": "00039464"},
    "Pabst": {"Torsten": "00021976"},
    "Pawlak": {"Bartosz": "00036381"},
    "Piepke": {"Torsten": "00021390"},
    "Plinke": {"Killian": "00044137"},
    "Pogodski": {"Enrico": "00046668"},
    "Quint": {"Stefan": "00035718"},
    "Rimba": {"Rimba Gona": "00046108"},
    "Sarwatka": {"Heiko": "00028747"},
    "Scheil": {"Eric-Rene": "00038579", "Rene": "00020851"},
    "Schlichting": {"Michael": "00021452"},
    "Schlutt": {"Hubert": "00020880", "Rene": "00042932"},
    "Schmieder": {"Steffen": "00046286"},
    "Schneider": {"Matthias": "00045495"},
    "Schulz": {"Julian": "00049130", "Stephan": "00041558"},
    "Singh": {"Jagtar": "00040902"},
    "Stoltz": {"Thorben": "00040991"},
    "Thal": {"Jannic": "00046006"},
    "Tumanow": {"Vasilli": "00045019"},
    "Wachnowski": {"Klaus": "00026019"},
    "Wendel": {"Danilo": "00048994"},
    "Wille": {"Rene": "00021393"},
    "Wisniewski": {"Krzysztof": "00046550"},
    "Zander": {"Jan": "00042454"},
    "Zosel": {"Ingo": "00026303"},
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
