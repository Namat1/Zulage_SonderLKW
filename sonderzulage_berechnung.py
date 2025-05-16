
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl import Workbook
import calendar

# Deutsche Monatsnamen
german_months = ["Dummy", "Januar", "Februar", "März", "April", "Mai", "Juni",
                 "Juli", "August", "September", "Oktober", "November", "Dezember"]

def get_german_month_name(month_number):
    return german_months[month_number]

def define_art(value):
    if value in [602, 156]:
        return "Gigaliner"
    elif value in [350, 620]:
        return "Tandem"
    elif value in [520, 266]:
        return "Gliederzug"
    return "Unbekannt"

def calculate_earnings(row):
    try:
        if pd.notnull(row["LKW"]) and "-" in row["LKW"]:
            nummer = int(row["LKW"].split("-")[1])
            if nummer in [266, 520, 620, 350]:
                return 20
            elif nummer in [602, 156]:
                return 40
    except:
        pass
    return 0

def format_date_with_german_weekday(date):
    mapping = {
        "Monday": "Montag", "Tuesday": "Dienstag", "Wednesday": "Mittwoch",
        "Thursday": "Donnerstag", "Friday": "Freitag", "Saturday": "Samstag", "Sunday": "Sonntag"
    }
    day_name = date.strftime("%A")
    german_day = mapping.get(day_name, day_name)
    kw = int(date.strftime("%W")) + 1
    return date.strftime(f"%d.%m.%Y ({german_day}, KW{kw})")

def apply_styles(sheet):
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    total_fill = PatternFill("solid", fgColor="C7B7B3")
    data_fill = PatternFill("solid", fgColor="FFFFFF")
    name_fill = PatternFill("solid", fgColor="F2ECE8")
    first_block_fill = PatternFill("solid", fgColor="95b3d7")
    is_first_in_block = True

    for row in sheet.iter_rows(min_col=1, max_col=5, values_only=False):
        first_val = str(row[0].value).strip() if row[0].value else ""
        if "Gesamtverdienst" in first_val:
            for cell in row:
                cell.fill = total_fill
                cell.font = Font(bold=True, size=11)
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.border = thin_border
                if cell.column == 5:
                    cell.number_format = '#,##0.00 €'
            is_first_in_block = True
        elif is_first_in_block and first_val:
            for cell in row:
                cell.fill = first_block_fill
                cell.font = Font(bold=True, size=12, italic=True)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border
                if cell.column == 5:
                    cell.number_format = '#,##0.00 €'
            is_first_in_block = False
        elif first_val:
            for cell in row:
                cell.fill = name_fill
                cell.font = Font(bold=True, size=11)
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.border = thin_border
                if cell.column == 5:
                    cell.number_format = '#,##0.00 €'
        else:
            for cell in row:
                cell.fill = data_fill
                cell.font = Font(size=11)
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.border = thin_border
                if cell.column == 5:
                    try:
                        cell.value = float(cell.value)
                        cell.number_format = '#,##0.00 €'
                    except:
                        pass
            if not first_val:
                is_first_in_block = True

    for col in sheet.columns:
        max_len = max(len(str(cell.value) or "") for cell in col)
        sheet.column_dimensions[get_column_letter(col[0].column)].width = max_len + 3

    sheet.row_dimensions[1].hidden = True

def add_summary(sheet, summary_data, start_col=9, month_name=""):
    header_fill = PatternFill("solid", fgColor="95b3d7")
    name_fill = PatternFill("solid", fgColor="F2ECE8")
    total_fill = PatternFill("solid", fgColor="C7B7B3")
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))

    sheet.cell(row=2, column=start_col, value="Auszahlung Monat:").font = Font(bold=True)
    sheet.cell(row=2, column=start_col).fill = header_fill
    sheet.cell(row=2, column=start_col).border = border
    sheet.cell(row=2, column=start_col).alignment = Alignment(horizontal="center")

    sheet.cell(row=2, column=start_col + 1, value=month_name).font = Font(bold=True)
    sheet.cell(row=2, column=start_col + 1).fill = header_fill
    sheet.cell(row=2, column=start_col + 1).border = border
    sheet.cell(row=2, column=start_col + 1).alignment = Alignment(horizontal="right")

    headers = ["Name", "Personalnummer", "Gesamtverdienst (€)"]
    for i, header in enumerate(headers):
        cell = sheet.cell(row=3, column=start_col + i, value=header)
        cell.font = Font(bold=True)
        cell.fill = total_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = border

    summary_data.sort(key=lambda x: x[2], reverse=True)
    for i, (name, pnr, total) in enumerate(summary_data, start=4):
        sheet.cell(row=i, column=start_col, value=name).fill = name_fill
        sheet.cell(row=i, column=start_col + 1, value=pnr).fill = name_fill
        cell = sheet.cell(row=i, column=start_col + 2, value=total)
        cell.fill = name_fill
        cell.number_format = '#,##0.00 €'
        for j in range(3):
            sheet.cell(row=i, column=start_col + j).border = border

    last_row = len(summary_data) + 4
    sheet.cell(row=last_row, column=start_col, value="Gesamtsumme").font = Font(bold=True)
    summe = sum(x[2] for x in summary_data)
    sheet.cell(row=last_row, column=start_col + 2, value=summe).number_format = '#,##0.00 €'
    for j in range(3):
        cell = sheet.cell(row=last_row, column=start_col + j)
        cell.fill = total_fill
        cell.font = Font(bold=True)
        cell.border = border

# Personalnummern-Zuordnung (gekürzt für Beispiel)
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
    "Paul": {"Toralf": "00010490"},
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

def main():
    st.title("Zulage - Sonderfahrzeuge - Ab 2025")

    uploaded_files = st.file_uploader("Lade eine oder mehrere Excel-Dateien hoch", type=["xlsx", "xls"], accept_multiple_files=True)

    if uploaded_files:
        all_data = pd.DataFrame()

        for uploaded_file in uploaded_files:
            try:
                df = pd.read_excel(uploaded_file, sheet_name="Touren", header=0)
                filtered_df = df[df.iloc[:, 13].astype(str).str.contains("AZ", case=False, na=False)]
                filtered_df["Datum"] = pd.to_datetime(filtered_df.iloc[:, 14], format="%d.%m.%Y", errors="coerce")
                filtered_df = filtered_df[filtered_df["Datum"] >= pd.Timestamp("2025-01-01")]

                if filtered_df.empty:
                    st.warning(f"Keine passenden Daten in der Datei {uploaded_file.name} gefunden.")
                    continue

                cols = [0, 3, 4, 10, 11, 12, 14]
                data = filtered_df.iloc[:, cols]
                data.columns = ["Tour", "Nachname", "Vorname", "LKW1", "LKW", "Art", "Datum"]
                data["LKW"] = data["LKW"].apply(lambda x: f"E-{x}" if pd.notnull(x) else x)
                data["Art"] = data["LKW"].apply(lambda x: define_art(int(x.split("-")[1])) if pd.notnull(x) and "-" in x else "Unbekannt")
                data["Verdienst"] = data.apply(calculate_earnings, axis=1)
                data["Monat"] = data["Datum"].dt.month
                data["Jahr"] = data["Datum"].dt.year
                all_data = pd.concat([all_data, data], ignore_index=True)

            except Exception as e:
                st.error(f"Fehler beim Verarbeiten von {uploaded_file.name}: {e}")

        if not all_data.empty:
            output_file = "zulagen_auswertung.xlsx"
            with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                for (jahr, monat), df_month in all_data.groupby(["Jahr", "Monat"]):
                    sheet_name = f"{get_german_month_name(monat)} {jahr}"
                    sheet_data = []
                    summary = []

                    for (nachname, vorname), gruppe in df_month.groupby(["Nachname", "Vorname"]):
                        name = f"{vorname} {nachname}"
                        summe = gruppe["Verdienst"].sum()
                        pnr = name_to_personalnummer.get(nachname, {}).get(vorname, "Unbekannt")
                        summary.append([name, pnr, summe])
                        sheet_data.append([name, "", "", "", ""])
                        sheet_data.append(["Datum", "Tour", "LKW", "Art", "Verdienst"])
                        for _, zeile in gruppe.iterrows():
                            datum = format_date_with_german_weekday(zeile["Datum"])
                            sheet_data.append([datum, zeile["Tour"], zeile["LKW"], zeile["Art"], zeile["Verdienst"]])
                        sheet_data.append(["Gesamtverdienst", "", "", "", summe])
                        sheet_data.append([])

                    df_export = pd.DataFrame(sheet_data)
                    df_export.to_excel(writer, sheet_name=sheet_name[:31], index=False)
                    ws = writer.sheets[sheet_name[:31]]
                    add_summary(ws, summary, start_col=9, month_name=sheet_name)
                    apply_styles(ws)

            with open(output_file, "rb") as f:
                st.download_button("Download Auswertung", f, file_name=output_file)

if __name__ == "__main__":
    main()
