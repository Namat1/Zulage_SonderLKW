import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

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
            if nummer in [602, 156]:
                return 40
            elif nummer in [620, 350, 520, 266]:
                return 20
    except:
        pass
    return 0

def format_date_with_german_weekday(date):
    wochentage_mapping = {
        "Monday": "Montag", "Tuesday": "Dienstag", "Wednesday": "Mittwoch",
        "Thursday": "Donnerstag", "Friday": "Freitag", "Saturday": "Samstag", "Sunday": "Sonntag"
    }
    german_weekday = wochentage_mapping.get(date.strftime("%A"), "")
    kw = int(date.strftime("%W")) + 1
    return date.strftime(f"%d.%m.%Y ({german_weekday}, KW{kw})")

def apply_styles(sheet):
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))
    total_fill = PatternFill(start_color="C7B7B3", end_color="C7B7B3", fill_type="solid")
    data_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    name_fill = PatternFill(start_color="F2ECE8", end_color="F2ECE8", fill_type="solid")
    first_block_fill = PatternFill(start_color="95b3d7", end_color="95b3d7", fill_type="solid")
    is_first_in_block = True

    for row in sheet.iter_rows(min_col=1, max_col=5, values_only=False):
        first_cell_value = str(row[0].value).strip() if row[0].value else ""
        if "Gesamtverdienst" in first_cell_value:
            for cell in row:
                cell.fill = total_fill
                cell.font = Font(bold=True, size=11)
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.border = thin_border
                if cell.column == 5:
                    cell.number_format = '#,##0.00 €'
            is_first_in_block = True
        elif is_first_in_block and first_cell_value:
            for cell in row:
                cell.fill = first_block_fill
                cell.font = Font(bold=True, size=12, italic=True)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border
                if cell.column == 5:
                    cell.number_format = '#,##0.00 €'
            is_first_in_block = False
        elif first_cell_value:
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
            if not first_cell_value:
                is_first_in_block = True

    for col in sheet.columns:
        max_length = max(len(str(cell.value) or "") for cell in col)
        col_letter = get_column_letter(col[0].column)
        sheet.column_dimensions[col_letter].width = max_length + 3

    sheet.row_dimensions[1].hidden = True

def add_summary(sheet, summary_data, start_col=9, month_name=""):
    header_fill = PatternFill(start_color="95b3d7", end_color="95b3d7", fill_type="solid")
    name_fill = PatternFill(start_color="F2ECE8", end_color="F2ECE8", fill_type="solid")
    verdienst_fill = name_fill
    total_fill = PatternFill(start_color="C7B7B3", end_color="C7B7B3", fill_type="solid")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                         top=Side(style='thin'), bottom=Side(style='thin'))

    sheet.cell(row=2, column=start_col, value="Auszahlung Monat:").fill = header_fill
    sheet.cell(row=2, column=start_col + 1, value=month_name).fill = header_fill

    headers = ["Name", "Personalnummer", "Gesamtverdienst (€)"]
    for idx, header in enumerate(headers, start=start_col):
        cell = sheet.cell(row=3, column=idx, value=header)
        cell.font = Font(bold=True, size=12)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = header_fill
        cell.border = thin_border

    summary_data.sort(key=lambda x: x[2], reverse=True)

    for i, (name, personalnummer, total) in enumerate(summary_data, start=4):
        sheet.cell(row=i, column=start_col, value=name).fill = name_fill
        pn_cell = sheet.cell(row=i, column=start_col + 1, value=personalnummer)
        pn_cell.number_format = '00000000'
        pn_cell.fill = name_fill
        total_cell = sheet.cell(row=i, column=start_col + 2, value=total)
        total_cell.number_format = '#,##0.00 €'
        total_cell.fill = verdienst_fill
        for col in range(start_col, start_col + 3):
            cell = sheet.cell(row=i, column=col)
            cell.font = Font(bold=True, size=12)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border

    total_row = len(summary_data) + 4
    sheet.cell(row=total_row, column=start_col, value="Gesamtsumme").fill = total_fill
    sheet.cell(row=total_row, column=start_col + 2, value=sum(x[2] for x in summary_data)).number_format = '#,##0.00 €'
    for col in range(start_col, start_col + 3):
        cell = sheet.cell(row=total_row, column=col)
        cell.font = Font(bold=True, size=12)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        if cell.fill.fill_type is None:
            cell.fill = total_fill

# Hauptfunktion
def main():
    st.title("Zulage - Sonderfahrzeuge - Ab 2025")
    uploaded_files = st.file_uploader("Lade eine oder mehrere Excel-Dateien hoch", type=["xlsx", "xls"], accept_multiple_files=True)

    if uploaded_files:
        all_data = pd.DataFrame()
        for uploaded_file in uploaded_files:
            try:
                df = pd.read_excel(uploaded_file, sheet_name="Touren", header=0)
                df.columns = [f"Spalte_{i}" for i in range(len(df.columns))]
                filtered_df = df[df["Spalte_13"].astype(str).str.upper().str.contains("AZ", na=False)]

                if filtered_df.empty:
                    st.warning(f"Keine AZ-Zeilen in Datei {uploaded_file.name}")
                    continue

                extracted_data = pd.DataFrame()
                extracted_data["Nachname"] = filtered_df["Spalte_3"]
                extracted_data["Vorname"] = filtered_df["Spalte_5"]
                extracted_data["LKW"] = filtered_df["Spalte_11"]
                extracted_data["Datum"] = pd.to_datetime(filtered_df["Spalte_14"], errors="coerce")
                extracted_data["Tour"] = filtered_df["Spalte_17"] if "Spalte_17" in filtered_df.columns else ""

                extracted_data["LKW"] = extracted_data["LKW"].apply(lambda x: f"E-{int(x)}" if pd.notnull(x) else x)
                extracted_data["Art"] = extracted_data["LKW"].apply(
                    lambda x: define_art(int(x.split("-")[1])) if pd.notnull(x) and "-" in x else "Unbekannt"
                )
                extracted_data["Verdienst"] = extracted_data.apply(calculate_earnings, axis=1)
                extracted_data["Monat"] = extracted_data["Datum"].dt.month
                extracted_data["Jahr"] = extracted_data["Datum"].dt.year

                all_data = pd.concat([all_data, extracted_data], ignore_index=True)
            except Exception as e:
                st.error(f"Fehler beim Einlesen von {uploaded_file.name}: {e}")

        if not all_data.empty:
            output_file = "Zulage_Sonderfahrzeuge_2025.xlsx"
            try:
                with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                    fallback = True
                    sorted_data = all_data.sort_values(by=["Jahr", "Monat"])
                    for year, month in sorted_data[["Jahr", "Monat"]].drop_duplicates().values:
                        month_data = sorted_data[(sorted_data["Monat"] == month) & (sorted_data["Jahr"] == year)]
                        if not month_data.empty:
                            fallback = False
                            sheet_name = f"{get_german_month_name(month)} {year}"
                            sheet_data = []
                            summary_data = []

                            for (nachname, vorname), group in month_data.groupby(["Nachname", "Vorname"]):
                                total_earnings = group["Verdienst"].sum()
                                personalnummer = name_to_personalnummer.get(nachname, {}).get(vorname, "Unbekannt")
                                summary_data.append([f"{vorname} {nachname}", personalnummer, total_earnings])
                                sheet_data.append([f"{vorname} {nachname}", "", "", "", ""])
                                sheet_data.append(["Datum", "Tour", "LKW", "Art", "Verdienst"])
                                for _, row in group.iterrows():
                                    sheet_data.append([
                                        format_date_with_german_weekday(row["Datum"]),
                                        row["Tour"],
                                        row["LKW"],
                                        row["Art"],
                                        row["Verdienst"]
                                    ])
                                sheet_data.append(["Gesamtverdienst", "", "", "", total_earnings])
                                sheet_data.append([])

                            pd.DataFrame(sheet_data).to_excel(writer, index=False, sheet_name=sheet_name[:31])
                            sheet = writer.sheets[sheet_name[:31]]
                            add_summary(sheet, summary_data, start_col=9, month_name=sheet_name)
                            apply_styles(sheet)

                    if fallback:
                        pd.DataFrame([["Keine gültigen AZ-Zeilen vorhanden."]]).to_excel(writer, sheet_name="Hinweis", index=False)

                with open(output_file, "rb") as file:
                    st.download_button(
                        label="Download Auswertung",
                        data=file.read(),
                        file_name=output_file,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            except Exception as e:
                st.error(f"Fehler beim Exportieren: {e}")

if __name__ == "__main__":
    main()
