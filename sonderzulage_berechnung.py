import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# Personalnummer-Zuordnung (gekürzt für Übersicht)
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

def get_german_month_name(month_number):
    german_months = ["Dummy", "Januar", "Februar", "März", "April", "Mai", "Juni",
                     "Juli", "August", "September", "Oktober", "November", "Dezember"]
    return german_months[month_number]

def define_art(value):
    if value in [602, 156]: return "Gigaliner"
    elif value in [350, 620]: return "Tandem"
    elif value in [520, 266]: return "Gliederzug"
    return "Unbekannt"

def format_date_with_german_weekday(date):
    days = {"Monday": "Montag", "Tuesday": "Dienstag", "Wednesday": "Mittwoch",
            "Thursday": "Donnerstag", "Friday": "Freitag", "Saturday": "Samstag", "Sunday": "Sonntag"}
    weekday = days.get(date.strftime("%A"), date.strftime("%A"))
    kw = int(date.strftime("%W")) + 1 if int(date.strftime("%W")) < 53 else 1
    return date.strftime(f"%d.%m.%Y ({weekday}, KW{kw})")

def calc_verdienst(row):
    v = 0
    lkw_vals = [
        row["LKW1"],
        row["LKW"].split("-")[1] if pd.notnull(row["LKW"]) and "-" in row["LKW"] else None,
        row["Art"]
    ]
    for val in lkw_vals:
        try:
            num = int(val) if isinstance(val, str) and val.isdigit() else val
            if num in [602, 156]: v += 40
            elif num in [620, 350, 520, 266]: v += 20
        except: continue

    # NEU: Füngers-Zulage nur wenn Name + Vorname vorhanden
    if (
        pd.notnull(row["Kommentar"]) and "füngers" in str(row["Kommentar"]).lower()
        and pd.notnull(row["Nachname"]) and pd.notnull(row["Vorname"])
    ):
        v += 20

    return v

def apply_styles(sheet):
    thin = Border(left=Side(style='thin'), right=Side(style='thin'),
                  top=Side(style='thin'), bottom=Side(style='thin'))
    fill_total = PatternFill("solid", fgColor="C7B7B3")
    fill_data = PatternFill("solid", fgColor="FFFFFF")
    fill_name = PatternFill("solid", fgColor="F2ECE8")
    fill_first = PatternFill("solid", fgColor="95b3d7")
    first_in_block = True

    for row in sheet.iter_rows(min_col=1, max_col=5):
        first_val = str(row[0].value).strip() if row[0].value else ""
        for cell in row:
            if "Gesamtverdienst" in first_val:
                cell.fill = fill_total
                cell.font = Font(bold=True, size=11)
                cell.alignment = Alignment("right", "center")
                cell.border = thin
                if cell.column == 5:
                    cell.number_format = '#,##0.00 €'
                first_in_block = True
            elif first_in_block and first_val:
                cell.fill = fill_first
                cell.font = Font(bold=True, size=12, italic=True)
                cell.alignment = Alignment("center", "center")
                cell.border = thin
                if cell.column == 5:
                    cell.number_format = '#,##0.00 €'
                first_in_block = False
            elif first_val:
                cell.fill = fill_name
                cell.font = Font(bold=True, size=11)
                cell.alignment = Alignment("right", "center")
                cell.border = thin
                if cell.column == 5:
                    cell.number_format = '#,##0.00 €'
            else:
                cell.fill = fill_data
                cell.font = Font(size=11)
                cell.alignment = Alignment("right", "center")
                cell.border = thin
                if cell.column == 5:
                    try:
                        cell.value = float(cell.value)
                        cell.number_format = '#,##0.00 €'
                    except:
                        pass
                if not first_val:
                    first_in_block = True

    for col in sheet.columns:
        max_len = max(len(str(c.value) or "") for c in col)
        col_letter = get_column_letter(col[0].column)
        sheet.column_dimensions[col_letter].width = max_len + 3

    sheet.row_dimensions[1].hidden = True

def main():
    st.title("Zulage - Sonderfahrzeuge - Ab 2025")
    uploaded_files = st.file_uploader("Lade eine oder mehrere Excel-Dateien hoch", type=["xlsx", "xls"], accept_multiple_files=True)

    if uploaded_files:
        all_data = pd.DataFrame()

        for file in uploaded_files:
            try:
                df = pd.read_excel(file, sheet_name="Touren", header=0)
                filtered = df.iloc[4:]
                filtered = filtered[pd.to_datetime(filtered.iloc[:, 14], errors="coerce").notna()]

                columns = [0, 3, 4, 10, 11, 12, 14, 15]
                extracted = filtered.iloc[:, columns]
                extracted.columns = ["Tour", "Nachname", "Vorname", "LKW1", "LKW", "Art", "Datum", "Kommentar"]
                extracted["LKW"] = extracted["LKW"].apply(lambda x: f"E-{x}" if pd.notnull(x) else x)
                extracted["Art"] = extracted["LKW"].apply(lambda x: define_art(int(x.split("-")[1])) if pd.notnull(x) and "-" in x else "Unbekannt")
                extracted["Datum"] = pd.to_datetime(extracted["Datum"], errors="coerce")
                extracted["Verdienst"] = extracted.apply(calc_verdienst, axis=1)
                extracted["Monat"] = extracted["Datum"].dt.month
                extracted["Jahr"] = extracted["Datum"].dt.year
                all_data = pd.concat([all_data, extracted], ignore_index=True)

            except Exception as e:
                st.error(f"Fehler beim Einlesen von {file.name}: {e}")

        if not all_data.empty:
            output_file = "Zulage_Sonderfahrzeuge_2025.xlsx"
            try:
                with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                    for (year, month), group in all_data.groupby(["Jahr", "Monat"]):
                        sheet_data, summary_data = [], []
                        for (nach, vor), df_group in group.groupby(["Nachname", "Vorname"]):
                            total = df_group["Verdienst"].sum()
                            personalnummer = name_to_personalnummer.get(nach, {}).get(vor, "Unbekannt")
                            summary_data.append([f"{vor} {nach}", personalnummer, total])
                            sheet_data.append([f"{vor} {nach}", "", "", "", ""])
                            sheet_data.append(["Datum", "Tour", "LKW", "Art", "Verdienst"])
                            for _, row in df_group.iterrows():
                                sheet_data.append([
                                    format_date_with_german_weekday(row["Datum"]),
                                    row["Tour"], row["LKW"], row["Art"], row["Verdienst"]
                                ])
                            sheet_data.append(["Gesamtverdienst", "", "", "", total])
                            sheet_data.append([])

                        df_sheet = pd.DataFrame(sheet_data)
                        sheet_name = f"{get_german_month_name(month)} {year}"
                        df_sheet.to_excel(writer, index=False, sheet_name=sheet_name[:31])
                        sheet = writer.sheets[sheet_name[:31]]
                        apply_styles(sheet)

                with open(output_file, "rb") as f:
                    st.download_button("Download Auswertung", f, file_name=output_file)

            except Exception as e:
                st.error(f"Fehler beim Exportieren: {e}")

if __name__ == "__main__":
    main()
