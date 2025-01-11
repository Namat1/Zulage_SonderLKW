import pandas as pd
import streamlit as st
import calendar
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


def apply_styles(sheet):
    """
    Dynamische Anwendung eines klaren und übersichtlichen Business-Stils:
    - Namenszeilen: Blau, fett.
    - Datum-/Kopfzeilen: Hellgrau, fett.
    - Datenzeilen: Weiß, standard.
    """
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    name_fill = PatternFill(start_color="D9EAF7", end_color="D9EAF7", fill_type="solid")  # Blau für Namenszeilen
    header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")  # Hellgrau für Kopfzeilen
    data_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")  # Weiß für Datenzeilen

    for row_idx, row in enumerate(sheet.iter_rows(), start=1):
        if row[0].value and row[0].value.split(" ")[0].isalpha():  # Namenszeilen
            for cell in row:
                cell.fill = name_fill
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center")
                cell.border = thin_border
        elif "Datum" in str(row[0].value):  # Kopfzeilen mit "Datum"
            for cell in row:
                cell.fill = header_fill
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center")
                cell.border = thin_border
        else:  # Datenzeilen
            for cell in row:
                cell.fill = data_fill
                cell.font = Font(bold=False)
                cell.alignment = Alignment(horizontal="left")
                cell.border = thin_border


def main():
    st.title("Touren-Auswertung mit klarer Trennung der Namenszeile")

    uploaded_files = st.file_uploader("Lade eine oder mehrere Excel-Dateien hoch", type=["xlsx", "xls"], accept_multiple_files=True)

    if uploaded_files:
        all_data = pd.DataFrame()  # DataFrame zur Speicherung aller Daten

        for uploaded_file in uploaded_files:
            try:
                df = pd.read_excel(uploaded_file, sheet_name="Touren", header=0)
                filtered_df = df[df.iloc[:, 13].str.contains(r'(?i)\b(AZ)\b', na=False)]
                if filtered_df.empty:
                    st.warning(f"Keine passenden Daten im Blatt 'Touren' der Datei {uploaded_file.name} gefunden.")
                    continue

                columns_to_extract = [0, 3, 4, 10, 11, 12, 14]
                extracted_data = filtered_df.iloc[:, columns_to_extract]
                extracted_data.columns = ["Tour", "Nachname", "Vorname", "LKW1", "LKW2", "LKW3", "Datum"]
                extracted_data["Datum"] = pd.to_datetime(extracted_data["Datum"], format="%d.%m.%Y", errors="coerce")

                def calculate_earnings(row):
                    lkw_values = [row["LKW1"], row["LKW2"], row["LKW3"]]
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
                            for (nachname, vorname), group in month_data.groupby(["Nachname", "Vorname"]):
                                sheet_data.append([f"{vorname} {nachname}", "", "", "", ""])
                                sheet_data.append(["Datum", "Tour", "LKW2", "LKW3", "Verdienst"])
                                for _, row in group.iterrows():
                                    sheet_data.append([
                                        row["Datum"].strftime("%d.%m.%Y"),
                                        row["Tour"],
                                        row["LKW2"],
                                        row["LKW3"],
                                        row["Verdienst"]
                                    ])
                                total_earnings = group["Verdienst"].sum()
                                sheet_data.append(["Gesamtverdienst", "", "", "", total_earnings])
                                sheet_data.append([])

                            sheet_df = pd.DataFrame(sheet_data)
                            sheet_df.to_excel(writer, index=False, sheet_name=month_name[:31])

                    workbook = writer.book
                    for sheet_name in workbook.sheetnames:
                        sheet = workbook[sheet_name]
                        for col in sheet.columns:
                            max_length = max(len(str(cell.value)) for cell in col if cell.value)
                            col_letter = get_column_letter(col[0].column)
                            sheet.column_dimensions[col_letter].width = max_length + 2
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
