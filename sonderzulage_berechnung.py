import pandas as pd
import streamlit as st
import calendar
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import calendar

# Deutsche Monatsnamen
german_months = [
    "Dummy",  # Platzhalter für den Index 0
    "Januar", "Februar", "März", "April", "Mai", "Juni",
    "Juli", "August", "September", "Oktober", "November", "Dezember"
]

# Funktion für deutschen Monatsnamen
def get_german_month_name(month_number):
    return german_months[month_number]


# Personalnummer-Zuordnung (verkürzt)
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
    Optische Formatierung der Excel-Daten, einschließlich separater Formatierung für die erste Zeile jedes Blocks.
    """
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    total_fill = PatternFill(start_color="C7B7B3", end_color="C7B7B3", fill_type="solid")
    data_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    name_fill = PatternFill(start_color="F2ECE8", end_color="F2ECE8", fill_type="solid")
    first_block_fill = PatternFill(start_color="7382AD", end_color="7382AD", fill_type="solid")  # Separates Styling für erste Zeilen

    is_first_in_block = True  # Markiert die erste Zeile eines Blocks

    for row_idx, row in enumerate(sheet.iter_rows(min_col=1, max_col=5), start=1):
        first_cell_value = str(row[0].value).strip() if row[0].value else ""

        # Gesamtverdienst-Zeilen formatieren
        if "Gesamtverdienst" in first_cell_value:
            for cell in row:
                cell.fill = total_fill
                cell.font = Font(bold=True, size=11)
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.border = thin_border
                if cell.column == 5:  # Euro-Format für Gesamtverdienst
                    cell.number_format = '#,##0.00 €'
            is_first_in_block = True  # Start eines neuen Blocks

        # Erste Zeile jedes Blocks formatieren
        elif is_first_in_block and first_cell_value:
            for cell in row:
                cell.fill = first_block_fill  # Spezielles Styling für erste Zeile
                cell.font = Font(bold=True, size=12, italic=True)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border
                if cell.column == 5:  # Euro-Format für Gesamtverdienst
                    cell.number_format = '#,##0.00 €'
            is_first_in_block = False  # Nach der ersten Zeile

        # Übrige Namen-Zeilen formatieren
        elif first_cell_value:
            for cell in row:
                cell.fill = name_fill
                cell.font = Font(bold=True, size=11)
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.border = thin_border
                if cell.column == 5:  # Euro-Format für Gesamtverdienst
                    cell.number_format = '#,##0.00 €'

        # Datenzeilen formatieren
        else:
            for cell in row:
                cell.fill = data_fill
                cell.font = Font(size=11)
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.border = thin_border

                # Setze das Euro-Zeichen für die Spalte "Verdienst" (Spalte 5)
                if cell.column == 5:
                    try:
                        cell.value = float(cell.value)  # Konvertiere den Wert in einen numerischen Typ
                        cell.number_format = '#,##0.00 €'
                    except (ValueError, TypeError):
                        pass  # Ignoriere nicht-numerische Werte

            # Markiere die nächste Zeile als potenziell erster Block
            if not first_cell_value:
                is_first_in_block = True

    # Spaltenbreiten automatisch anpassen
    for col in sheet.columns:
        max_length = max(len(str(cell.value) or "") for cell in col)
        col_letter = get_column_letter(col[0].column)
        sheet.column_dimensions[col_letter].width = max_length + 3




 # Erste Zeile ausblenden
    sheet.row_dimensions[1].hidden = True


def format_date_with_german_weekday(date):
    wochentage_mapping = {
        "Monday": "Montag",
        "Tuesday": "Dienstag",
        "Wednesday": "Mittwoch",
        "Thursday": "Donnerstag",
        "Friday": "Freitag",
        "Saturday": "Samstag",
        "Sunday": "Sonntag"
    }
    english_weekday = date.strftime("%A")  # Englischer Wochentag
    german_weekday = wochentage_mapping.get(english_weekday, english_weekday)  # Übersetzung
    return date.strftime(f"%d.%m.%Y ({german_weekday}, KW%W)")


def add_summary(sheet, summary_data, start_col=9, month_name=""):
    """
    Fügt eine farblich formatierte Zusammenfassung der Daten hinzu.
    """
    # Farben für Kopfzeile und Spalten
    header_fill = PatternFill(start_color="7382AD", end_color="7382AD", fill_type="solid")
    name_fill = PatternFill(start_color="FFEEE1", end_color="FFEEE1", fill_type="solid")
    personalnummer_fill = PatternFill(start_color="E1F7FF", end_color="E1F7FF", fill_type="solid")
    verdienst_fill = PatternFill(start_color="E1FFE7", end_color="E1FFE7", fill_type="solid")
    total_fill = PatternFill(start_color="C7B7B3", end_color="C7B7B3", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    # Monatsname in Zeile 2
    auszahlung_text = f"Auszahlung Monat: {month_name}" if month_name else "Auszahlung Monat: Unbekannt"
    auszahlung_cell = sheet.cell(row=2, column=start_col, value=auszahlung_text)
    sheet.merge_cells(start_row=2, start_column=start_col, end_row=2, end_column=start_col + 2)
    auszahlung_cell.font = Font(bold=True, size=12)
    auszahlung_cell.alignment = Alignment(horizontal="center", vertical="center")
    auszahlung_cell.fill = header_fill
    auszahlung_cell.border = thin_border

    # Zusammenfassungskopfzeilen
    for idx, header in enumerate(["Name", "Personalnummer", "Gesamtverdienst (€)"], start=start_col):
        cell = sheet.cell(row=3, column=idx)
        cell.value = header
        cell.font = Font(bold=True, size=12)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    # Einfügen der Daten mit Farbfüllung
    for i, (name, personalnummer, total) in enumerate(summary_data, start=4):
        name_cell = sheet.cell(row=i, column=start_col, value=name)
        name_cell.font = Font(bold=True, size=12)
        name_cell.fill = name_fill
        name_cell.border = thin_border

    # Personalnummer als Text schreiben
    personalnummer_cell = sheet.cell(row=i, column=start_col + 1, value=personalnummer)
    personalnummer_cell.font = Font(bold=True, size=12)
    personalnummer_cell.fill = personalnummer_fill
    personalnummer_cell.border = thin_border
    personalnummer_cell.number_format = '@'  # Textformatierung

    total_cell = sheet.cell(row=i, column=start_col + 2, value=total)
    total_cell.font = Font(bold=True, size=12)
    total_cell.fill = verdienst_fill
    total_cell.number_format = '#,##0.00 €'
    total_cell.border = thin_border




    # Gesamtsumme aller Verdienste
    total_row = len(summary_data) + 4
    sheet.cell(row=total_row, column=start_col, value="Gesamtsumme").font = Font(bold=True, size=12)
    total_sum_cell = sheet.cell(row=total_row, column=start_col + 2, value=sum(x[2] for x in summary_data))
    total_sum_cell.font = Font(bold=True, size=12)
    total_sum_cell.fill = total_fill
    total_sum_cell.number_format = '#,##0.00 €'
    total_sum_cell.border = thin_border

    # Leere Zellen mit Rahmen versehen
    for row in range(4, total_row + 1):
        for col in range(start_col, start_col + 3):
            if sheet.cell(row=row, column=col).value is None:
                sheet.cell(row=row, column=col).border = thin_border



def main():
    st.title("Zulage - Sonderfahrzeuge - Ab 2025")

    uploaded_files = st.file_uploader("Lade eine oder mehrere Excel-Dateien hoch", type=["xlsx", "xls"], accept_multiple_files=True)

    if uploaded_files:
        all_data = pd.DataFrame()

        for uploaded_file in uploaded_files:
            try:
                df = pd.read_excel(uploaded_file, sheet_name="Touren", header=0)
                filtered_df = df[df.iloc[:, 13].str.contains(r'(?i)\b(AZ)\b', na=False)]
                if not filtered_df.empty:
                    filtered_df["Datum"] = pd.to_datetime(filtered_df.iloc[:, 14], format="%d.%m.%Y", errors="coerce")
                    filtered_df = filtered_df[filtered_df["Datum"] >= pd.Timestamp("2025-01-01")]

                if filtered_df.empty:
                    st.warning(f"Keine passenden Daten in der Datei {uploaded_file.name} gefunden.")
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
                            sheet_name = f"{get_german_month_name(month)} {year}"  # Deutscher Monatsname
                            sheet_data = []
                            summary_data = []

                            for (nachname, vorname), group in month_data.groupby(["Nachname", "Vorname"]):
                                total_earnings = group["Verdienst"].sum()
                                personalnummer = name_to_personalnummer.get(nachname, {}).get(vorname, "Unbekannt")
                                summary_data.append([f"{vorname} {nachname}", personalnummer, total_earnings])

                                # Gruppendaten zusammenstellen
                                sheet_data.append([f"{vorname} {nachname}", "", "", "", ""])
                                sheet_data.append(["Datum", "Tour", "LKW", "Art", "Verdienst"])
                                for _, row in group.iterrows():
                                    formatted_date = format_date_with_german_weekday(row["Datum"])
                                    sheet_data.append([
                                        formatted_date,
                                        row["Tour"],
                                        row["LKW"],
                                        row["Art"],
                                        row["Verdienst"]
                                    ])

                                sheet_data.append(["Gesamtverdienst", "", "", "", total_earnings])
                                sheet_data.append([])

                            # Daten in Excel-Tabellenblatt schreiben
                            sheet_df = pd.DataFrame(sheet_data)
                            sheet_df.to_excel(writer, index=False, sheet_name=sheet_name[:31])  # Tabellenblatt mit deutschem Namen
                            sheet = writer.sheets[sheet_name[:31]]

                            # Zusammenfassung und Styling anwenden
                            add_summary(sheet, summary_data, start_col=9, month_name=sheet_name)
                            apply_styles(sheet)

                with open(output_file, "rb") as file:
                    st.download_button(
                        label="Download Auswertung",
                        data=file,
                        file_name="Zulage_Sonderfahrzeuge_2025.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"Fehler beim Exportieren der Datei: {e}")


if __name__ == "__main__":
    main()
