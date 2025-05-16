
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl import Workbook
import calendar

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

def define_art(value):
    if value in [602, 156]:
        return "Gigaliner"
    elif value in [350, 620]:
        return "Tandem"
    elif value in [520, 266]:
        return "Gliederzug"
    return "Unbekannt"

def get_german_month_name(month_number):
    german_months = ["Dummy", "Januar", "Februar", "März", "April", "Mai", "Juni",
                     "Juli", "August", "September", "Oktober", "November", "Dezember"]
    return german_months[month_number]

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

def main():
    st.title("Zulage - Sonderfahrzeuge - Ab 2025")

    uploaded_files = st.file_uploader("Lade eine oder mehrere Excel-Dateien hoch", type=["xlsx", "xls"], accept_multiple_files=True)

    if uploaded_files:
        all_data = pd.DataFrame()

        for uploaded_file in uploaded_files:
            try:
                df = pd.read_excel(uploaded_file, sheet_name="Touren", header=1)

                filtered_df = df[df["Bemerkungen"].astype(str).str.contains("AZ", case=False, na=False)]
                filtered_df["Datum"] = pd.to_datetime(filtered_df["Datum"], format="%d.%m.%Y", errors="coerce")
                filtered_df = filtered_df[filtered_df["Datum"] >= pd.Timestamp("2025-01-01")]

                if filtered_df.empty:
                    st.warning(f"Keine passenden Daten in der Datei {uploaded_file.name} gefunden.")
                    continue

                cols = ["Tour", "Name", "V-Name", "LKW1", "LKW", "Frz. Gr.", "Datum"]
                data = filtered_df[cols].copy()
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
                        pnr = "Unbekannt"
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
