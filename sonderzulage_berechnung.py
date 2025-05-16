import pandas as pd
import streamlit as st
import re
from datetime import datetime
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl import Workbook

def extract_lkw_nummer(lkw_text):
    match = re.search(r"E-(\d+)", str(lkw_text))
    return int(match.group(1)) if match else None

def calculate_earnings(row):
    nummer = extract_lkw_nummer(row["LKW"])
    if nummer in [266, 520, 620, 350]:
        return 20
    elif nummer in [602, 156]:
        return 40
    return 0

def get_german_month_name(month_number):
    german_months = ["Dummy", "Januar", "Februar", "März", "April", "Mai", "Juni",
                     "Juli", "August", "September", "Oktober", "November", "Dezember"]
    return german_months[month_number]

def format_date_with_german_weekday(date):
    mapping = {
        "Monday": "Montag", "Tuesday": "Dienstag", "Wednesday": "Mittwoch",
        "Thursday": "Donnerstag", "Friday": "Freitag",
        "Saturday": "Samstag", "Sunday": "Sonntag"
    }
    day_name = date.strftime("%A")
    german_day = mapping.get(day_name, day_name)
    kw = int(date.strftime("%W")) + 1
    return date.strftime(f"%d.%m.%Y ({german_day}, KW{kw})")

def apply_styles(sheet):
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))
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

def main():
    st.title("Zulage - Sonderfahrzeuge - Ab 2025")

    uploaded_files = st.file_uploader("Lade eine oder mehrere Excel-Dateien hoch", type=["xlsx"], accept_multiple_files=True)

    if uploaded_files:
        all_data = pd.DataFrame()

        for uploaded_file in uploaded_files:
            try:
                df = pd.read_excel(uploaded_file, sheet_name="Touren", header=None, skiprows=4)
                df = df.rename(columns={3: "Nachname", 4: "Vorname", 11: "LKW", 13: "AZ-Kennung", 14: "Datum", 0: "Tour"})
                df = df[["Tour", "Nachname", "Vorname", "LKW", "AZ-Kennung", "Datum"]]
                df = df[df["AZ-Kennung"].astype(str).str.contains("AZ", case=False, na=False)]
                df["Datum"] = pd.to_datetime(df["Datum"], errors="coerce")
                df = df[df["Datum"] >= pd.Timestamp("2025-01-01")]
                df["LKW"] = df["LKW"].apply(lambda x: f"E-{int(float(str(x).replace('E-', '').replace(',', '.')))}" if pd.notnull(x) else x)
                df["Verdienst"] = df.apply(calculate_earnings, axis=1)
                df["Monat"] = df["Datum"].dt.month
                df["Jahr"] = df["Datum"].dt.year
                all_data = pd.concat([all_data, df], ignore_index=True)
            except Exception as e:
                st.error(f"Fehler beim Einlesen der Datei {uploaded_file.name}: {e}")

        if not all_data.empty:
            output_file = "zulagen_exportiert.xlsx"
            with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                for (jahr, monat), monat_df in all_data.groupby(["Jahr", "Monat"]):
                    sheet_name = f"{get_german_month_name(monat)} {jahr}"
                    sheet_data = []
                    summary = []

                    for (nachname, vorname), gruppe in monat_df.groupby(["Nachname", "Vorname"]):
                        name = f"{vorname} {nachname}"
                        summe = gruppe["Verdienst"].sum()
                        summary.append([name, "Unbekannt", summe])
                        sheet_data.append([name, "", "", "", ""])
                        sheet_data.append(["Datum", "Tour", "LKW", "Art", "Verdienst"])
                        for _, zeile in gruppe.iterrows():
                            datum = format_date_with_german_weekday(zeile["Datum"])
                            art = "Gigaliner" if extract_lkw_nummer(zeile["LKW"]) in [156, 602] else "Tandem/Gliederzug" if extract_lkw_nummer(zeile["LKW"]) in [266, 350, 520, 620] else "Unbekannt"
                            sheet_data.append([datum, zeile["Tour"], zeile["LKW"], art, zeile["Verdienst"]])
                        sheet_data.append(["Gesamtverdienst", "", "", "", summe])
                        sheet_data.append([])

                    df_export = pd.DataFrame(sheet_data)
                    df_export.to_excel(writer, sheet_name=sheet_name[:31], index=False)
                    ws = writer.sheets[sheet_name[:31]]
                    apply_styles(ws)

            with open(output_file, "rb") as f:
                st.download_button("Download Excel-Auswertung", f, file_name=output_file)

if __name__ == "__main__":
    main()
