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
    # ... (alle weiteren aus deinem Mapping)
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

def apply_styles(sheet):
    thin = Border(left=Side(style='thin'), right=Side(style='thin'),
                  top=Side(style='thin'), bottom=Side(style='thin'))
    fill_total = PatternFill("solid", fgColor="C7B7B3")
    fill_data = PatternFill("solid", fgColor="FFFFFF")
    fill_name = PatternFill("solid", fgColor="F2ECE8")
    fill_first = PatternFill("solid", fgColor="95b3d7")
    first_in_block = True

    for row in sheet.iter_rows(min_col=1, max_col=5, values_only=False):
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
                filtered = df[df.iloc[:, 13].str.contains(r'(?i)\\b(AZ)\\b', na=False)]
                filtered["Datum"] = pd.to_datetime(filtered.iloc[:, 14], errors="coerce")
                filtered = filtered[filtered["Datum"] >= pd.Timestamp("2025-01-01")]
                if filtered.empty:
                    st.warning(f"Keine passenden Daten in der Datei {file.name} gefunden.")
                    continue

                columns = [0, 3, 4, 10, 11, 12, 14, 15]
                extracted = filtered.iloc[:, columns]
                extracted.columns = ["Tour", "Nachname", "Vorname", "LKW1", "LKW", "Art", "Datum", "Kommentar"]
                extracted["LKW"] = extracted["LKW"].apply(lambda x: f"E-{x}" if pd.notnull(x) else x)
                extracted["Art"] = extracted["LKW"].apply(lambda x: define_art(int(x.split("-")[1])) if pd.notnull(x) and "-" in x else "Unbekannt")

                def calc_verdienst(row):
                    v = 0
                    lkw_vals = [row["LKW1"], row["LKW"].split("-")[1] if pd.notnull(row["LKW"]) and "-" in row["LKW"] else None, row["Art"]]
                    for val in lkw_vals:
                        try:
                            num = int(val) if isinstance(val, str) and val.isdigit() else val
                            if num in [602, 156]: v += 40
                            elif num in [620, 350, 520, 266]: v += 20
                        except: continue
                    if pd.notnull(row["Kommentar"]) and "füngers" in str(row["Kommentar"]).lower():
                        v += 20
                    return v

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
                        name = f"{get_german_month_name(month)} {year}"
                        df_sheet.to_excel(writer, index=False, sheet_name=name[:31])
                        sheet = writer.sheets[name[:31]]
                        apply_styles(sheet)

                        from openpyxl import load_workbook
                        from openpyxl.worksheet.worksheet import Worksheet
                        add_summary(sheet, summary_data, start_col=9, month_name=name)

                with open(output_file, "rb") as f:
                    st.download_button("Download Auswertung", f, file_name=output_file)

            except Exception as e:
                st.error(f"Fehler beim Exportieren: {e}")

def add_summary(sheet, summary_data, start_col=9, month_name=""):
    fill = PatternFill("solid", fgColor="95b3d7")
    name_fill = PatternFill("solid", fgColor="F2ECE8")
    total_fill = PatternFill("solid", fgColor="C7B7B3")
    thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    sheet.cell(row=2, column=start_col, value="Auszahlung Monat:").font = Font(bold=True)
    sheet.cell(row=2, column=start_col + 1, value=month_name).font = Font(bold=True)

    headers = ["Name", "Personalnummer", "Gesamtverdienst (€)"]
    for i, head in enumerate(headers, start=start_col):
        cell = sheet.cell(row=3, column=i, value=head)
        cell.font = Font(bold=True)
        cell.fill = fill
        cell.border = thin

    summary_data.sort(key=lambda x: x[2], reverse=True)
    for i, (name, pers, total) in enumerate(summary_data, start=4):
        sheet.cell(row=i, column=start_col, value=name).fill = name_fill
        sheet.cell(row=i, column=start_col + 1, value=pers).fill = name_fill
        cell_total = sheet.cell(row=i, column=start_col + 2, value=total)
        cell_total.fill = name_fill
        cell_total.number_format = '#,##0.00 €'

    row_sum = len(summary_data) + 4
    sheet.cell(row=row_sum, column=start_col, value="Gesamtsumme").fill = total_fill
    sheet.cell(row=row_sum, column=start_col + 2, value=sum(x[2] for x in summary_data)).fill = total_fill
    sheet.cell(row=row_sum, column=start_col + 2).number_format = '#,##0.00 €'

if __name__ == "__main__":
    main()
