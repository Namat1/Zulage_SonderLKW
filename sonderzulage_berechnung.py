import pandas as pd
import streamlit as st
import re
import calendar
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# Deutsche Monatsnamen
german_months = [
    "Dummy",
    "Januar", "Februar", "März", "April", "Mai", "Juni",
    "Juli", "August", "September", "Oktober", "November", "Dezember"
]

def get_german_month_name(month_number):
    return german_months[month_number]

def define_art(value):
    if value in [602, 156]:
        return "Gigaliner"
    elif value in [350, 620]:
        return "Tandem"
    elif value == 520:
        return "Gliederzug"
    return "Unbekannt"

# (name_to_personalnummer wird hier vorausgesetzt, im Originalcode enthalten)
# (apply_styles und add_summary ebenfalls vorhanden und oben eingefügt)

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
    english_weekday = date.strftime("%A")
    german_weekday = wochentage_mapping.get(english_weekday, english_weekday)
    original_kw = int(date.strftime("%W"))
    adjusted_kw = original_kw + 1 if original_kw < 53 else 1
    return date.strftime(f"%d.%m.%Y ({german_weekday}, KW{adjusted_kw})")

def main():
    st.title("Zulage - Sonderfahrzeuge - Ab 2025")

    uploaded_files = st.file_uploader("Lade eine oder mehrere Excel-Dateien hoch", type=["xlsx", "xls"], accept_multiple_files=True)

    if uploaded_files:
        all_data = pd.DataFrame()

        for uploaded_file in uploaded_files:
            try:
                df = pd.read_excel(uploaded_file, sheet_name="Touren", header=0)
                filtered_df = df[df.iloc[:, 13].str.contains(r'(?i)\\b(AZ)\\b', na=False)]
                if not filtered_df.empty:
                    filtered_df["Datum"] = pd.to_datetime(filtered_df.iloc[:, 14], format="%d.%m.%Y", errors="coerce")
                    filtered_df = filtered_df[filtered_df["Datum"] >= pd.Timestamp("2025-01-01")]
                if filtered_df.empty:
                    st.warning(f"Keine passenden Daten in der Datei {uploaded_file.name} gefunden.")
                    continue

                columns_to_extract = [0, 3, 4, 10, 11, 12, 14]
                extracted_data = filtered_df.iloc[:, columns_to_extract]
                extracted_data.columns = ["Tour", "Nachname", "Vorname", "LKW1", "LKW", "Art", "Datum"]
                extracted_data["Kommentar"] = filtered_df.iloc[:, 15]
                extracted_data["LKW"] = extracted_data["LKW"].apply(lambda x: f"E-{x}" if pd.notnull(x) else x)
                extracted_data["Art"] = extracted_data["LKW"].apply(lambda x: define_art(int(x.split("-")[1])) if pd.notnull(x) and "-" in x else "Unbekannt")
                extracted_data["Datum"] = pd.to_datetime(extracted_data["Datum"], format="%d.%m.%Y", errors="coerce")
                extracted_data["Tour"] = extracted_data["Tour"].fillna(filtered_df.iloc[:, 16])

                def calculate_earnings(row):
                    earnings = 0
                    lkw_values = [
                        row["LKW1"],
                        row["LKW"].split("-")[1] if pd.notnull(row["LKW"]) and "-" in row["LKW"] else None,
                        row["Art"]
                    ]
                    for value in lkw_values:
                        try:
                            numeric_value = int(value) if isinstance(value, str) and value.isdigit() else value
                            if numeric_value in [602, 156]:
                                earnings += 40
                            elif numeric_value in [620, 350, 520]:
                                earnings += 20
                        except (ValueError, TypeError):
                            continue

                    try:
                        kommentar = str(row.get("Kommentar", "") or "")
                        if re.search(r"f[\üu]ngers?", kommentar, re.IGNORECASE) or re.search(r"a[-\\s]?haus", kommentar, re.IGNORECASE):
                            earnings += 20
                    except:
                        pass

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
                                    formatted_date = format_date_with_german_weekday(row["Datum"])
                                    sheet_data.append([
                                        formatted_date,
                                        row["Tour"],
                                        row["LKW"],
                                        row["Art"],
                                        row["Verdienst"]
                                    ])

                                    kommentar = str(row.get("Kommentar", "") or "")
                                    if re.search(r"f[\üu]ngers?", kommentar, re.IGNORECASE) or re.search(r"a[-\\s]?haus", kommentar, re.IGNORECASE):
                                        sheet_data.append([
                                            "",
                                            kommentar.strip(),
                                            "",
                                            "Zusatz (Füingers/Ahaus)",
                                            20
                                        ])

                                sheet_data.append(["Gesamtverdienst", "", "", "", total_earnings])
                                sheet_data.append([])

                            sheet_df = pd.DataFrame(sheet_data)
                            sheet_df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
                            sheet = writer.sheets[sheet_name[:31]]

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
