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
    Formatierung für die Hauptdaten im Sheet, begrenzt auf Spalten A bis E, 
    und automatische Anpassung der Spaltenbreite mit +1 für alle Spalten außer A.
    """
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    name_fill = PatternFill(start_color="D9EAF7", end_color="D9EAF7", fill_type="solid")
    header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    total_fill = PatternFill(start_color="DFF7DF", end_color="DFF7DF", fill_type="solid")
    data_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    for row_idx, row in enumerate(sheet.iter_rows(min_col=1, max_col=5), start=1):
        first_cell_value = str(row[0].value).strip() if row[0].value else ""

        if "Gesamtverdienst" in first_cell_value:
            for cell in row:
                cell.fill = total_fill
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="right")
                cell.border = thin_border
                if cell.column == 5 and isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0.00 €'

        elif first_cell_value and any(char.isalpha() for char in first_cell_value) and not "Datum" in first_cell_value:
            try:
                vorname, nachname = first_cell_value.split(" ", 1)
                vorname = "".join(vorname.strip().split()).title()
                nachname = "".join(nachname.strip().split()).title()
                personalnummer = (
                    name_to_personalnummer.get(nachname, {}).get(vorname)
                    or name_to_personalnummer.get(nachname, {}).get(vorname.replace("-", " "))
                    or name_to_personalnummer.get(nachname, {}).get(vorname.replace(" ", "-"))
                    or "Unbekannt"
                )
            except ValueError:
                personalnummer = "Unbekannt"

            sheet.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=5)
            row[0].value = f"{first_cell_value} - {personalnummer}"
            row[0].fill = name_fill
            row[0].font = Font(bold=True)
            row[0].alignment = Alignment(horizontal="center")
            for cell in row:
                cell.border = thin_border

        elif "Datum" in first_cell_value:
            for cell in row:
                cell.fill = header_fill
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="right")
                cell.border = thin_border

        else:
            for cell in row:
                cell.fill = data_fill
                cell.font = Font(bold=False)
                cell.alignment = Alignment(horizontal="right")
                cell.border = thin_border
                if cell.column == 5 and isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0.00 €'

    # Automatische Spaltenbreitenanpassung für alle Spalten
    for col in sheet.columns:
        max_length = max(len(str(cell.value) or "") for cell in col)
        col_letter = get_column_letter(col[0].column)
        if col_letter == "A":
            sheet.column_dimensions[col_letter].width = 17  # Feste Breite für Spalte A
        else:
            sheet.column_dimensions[col_letter].width = max_length + 1  # +1 für alle anderen Spalten

    # Erste Zeile ausblenden
    sheet.row_dimensions[1].hidden = True





def add_summary(sheet, summary_data, start_col=9):
    """
    Fügt eine Zusammenfassungstabelle (Name, Personalnummer, Gesamtverdienst) in das Sheet ein.
    Die Zusammenfassung wird formatiert, um klarer und übersichtlicher zu sein.
    """
    # Überschriften für die Zusammenfassung
    header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    sheet.cell(row=1, column=start_col, value="Zusammenfassung")
    sheet.cell(row=2, column=start_col, value="Name")
    sheet.cell(row=2, column=start_col + 1, value="Personalnummer")
    sheet.cell(row=2, column=start_col + 2, value="Gesamtverdienst (€)")

    # Formatierung der Kopfzeile
    for col in range(start_col, start_col + 3):
        header_cell = sheet.cell(row=2, column=col)
        header_cell.font = Font(bold=True)
        header_cell.fill = header_fill
        header_cell.alignment = Alignment(horizontal="center", vertical="center")
        header_cell.border = thin_border

    # Einfügen der Zusammenfassungsdaten
    for idx, (name, personalnummer, total) in enumerate(summary_data, start=3):
        sheet.cell(row=idx, column=start_col, value=name).border = thin_border
        sheet.cell(row=idx, column=start_col + 1, value=personalnummer).border = thin_border
        total_cell = sheet.cell(row=idx, column=start_col + 2, value=total)
        total_cell.border = thin_border
        total_cell.number_format = '#,##0.00 €'  # Währungsformat

    # Spaltenbreite automatisch anpassen
    for col in range(start_col, start_col + 3):
        max_length = max(
            len(str(sheet.cell(row=row, column=col).value) or "") for row in range(1, len(summary_data) + 3)
        )
        sheet.column_dimensions[get_column_letter(col)].width = max_length + 1


def main():
    st.title("Touren-Auswertung mit klarer Trennung der Namenszeile")

    uploaded_files = st.file_uploader("Lade eine oder mehrere Excel-Dateien hoch", type=["xlsx", "xls"], accept_multiple_files=True)

    if uploaded_files:
        all_data = pd.DataFrame()

        for uploaded_file in uploaded_files:
            try:
                df = pd.read_excel(uploaded_file, sheet_name="Touren", header=0)
                filtered_df = df[df.iloc[:, 13].str.contains(r'(?i)\b(AZ)\b', na=False)]
                if filtered_df.empty:
                    st.warning(f"Keine passenden Daten im Blatt 'Touren' der Datei {uploaded_file.name} gefunden.")
                    continue

                columns_to_extract = [0, 3, 4, 10, 11, 12, 14]
                extracted_data = filtered_df.iloc[:, columns_to_extract]
                extracted_data.columns = ["Tour", "Nachname", "Vorname", "LKW1", "LKW", "Art", "Datum"]
                extracted_data["Datum"] = pd.to_datetime(extracted_data["Datum"], format="%d.%m.%Y", errors="coerce")

                def calculate_earnings(row):
                    lkw_values = [row["LKW1"], row["LKW"], row["Art"]]
                    earnings = 0
                    for value in lkw_values:
                        if value in [602, 156]:
                            earnings += 40
                        elif value in [620, 350, 520]:
                            earnings += 20
                    return earnings

                extracted_data["Verdienst"] = extracted_data.apply(calculate_earnings, axis=1)
                extracted_data["Monat"] = extracted_data["Datum"].dt.month
                extracted_data["Jahr"] = extracted_data["Datum"].dt.year
                all_data = pd.concat([all_data, extracted_data], ignore_index=True)

            except Exception as e:
                st.error(f"Fehler beim Einlesen der Datei {uploaded_file.name}: {e}")

        if not all_data.empty:
            output_file = "touren_auswertung_korrekt.xlsx"
            try:
                with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                    sorted_data = all_data.sort_values(by=["Jahr", "Monat"])

                    for year, month in sorted_data[["Jahr", "Monat"]].drop_duplicates().values:
                        month_data = sorted_data[(sorted_data["Monat"] == month) & (sorted_data["Jahr"] == year)]
                        if not month_data.empty:
                            month_name_german = {
                                "January": "Januar", "February": "Februar", "March": "März", "April": "April",
                                "May": "Mai", "June": "Juni", "July": "Juli", "August": "August",
                                "September": "September", "October": "Oktober", "November": "November", "December": "Dezember"
                            }
                            month_name = f"{month_name_german[calendar.month_name[month]]} {year}"

                            sheet_data = []
                            summary_data = []
                            for (nachname, vorname), group in month_data.groupby(["Nachname", "Vorname"]):
                                total_earnings = group["Verdienst"].sum()
                                personalnummer = (
                                    name_to_personalnummer.get(nachname, {}).get(vorname, "Unbekannt")
                                )
                                summary_data.append([f"{vorname} {nachname}", personalnummer, total_earnings])

                                sheet_data.append([f"{vorname} {nachname}", "", "", "", ""])
                                sheet_data.append(["Datum", "Tour", "LKW", "Art", "Verdienst"])
                                for _, row in group.iterrows():
                                    sheet_data.append([
                                        row["Datum"].strftime("%d.%m.%Y"),
                                        row["Tour"],
                                        row["LKW"],
                                        row["Art"],
                                        row["Verdienst"]
                                    ])
                                sheet_data.append(["Gesamtverdienst", "", "", "", total_earnings])
                                sheet_data.append([])

                            sheet_df = pd.DataFrame(sheet_data)
                            sheet_df.to_excel(writer, index=False, sheet_name=month_name[:31])

                            sheet = writer.sheets[month_name[:31]]
                            add_summary(sheet, summary_data, start_col=9)

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
