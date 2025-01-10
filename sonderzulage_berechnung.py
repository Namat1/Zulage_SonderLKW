import pandas as pd
import streamlit as st
import calendar
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

def main():
    st.title("Professionelle Touren-Auswertung mit klarem Layout")

    # Mehrere Dateien hochladen
    uploaded_files = st.file_uploader("Lade eine oder mehrere Excel-Dateien hoch", type=["xlsx", "xls"], accept_multiple_files=True)

    if uploaded_files:
        all_data = pd.DataFrame()  # DataFrame zur Speicherung aller Daten

        for uploaded_file in uploaded_files:
            try:
                # Lesen des Blattes "Touren" aus der aktuellen Datei
                df = pd.read_excel(uploaded_file, sheet_name="Touren", header=0)
                st.write(f"Datei: {uploaded_file.name} - Blatt 'Touren' erfolgreich geladen.")

                # Filter auf Spalte 14 (Werte Az, AZ, az)
                try:
                    filtered_df = df[df.iloc[:, 13].str.contains(r'(?i)\b(AZ)\b', na=False)]
                    if filtered_df.empty:
                        st.warning(f"Keine passenden Daten im Blatt 'Touren' der Datei {uploaded_file.name} gefunden.")
                        continue
                except Exception as e:
                    st.error(f"Fehler beim Filtern der Daten in Datei {uploaded_file.name}: {e}")
                    continue

                # Relevante Spalten extrahieren
                columns_to_extract = [0, 3, 4, 10, 11, 12, 14]  # Spalten: 1, 4, 5, 11, 12, 14, 15
                try:
                    extracted_data = filtered_df.iloc[:, columns_to_extract]
                    extracted_data.columns = ["Tour", "Nachname", "Vorname", "LKW1", "LKW2", "LKW3", "Datum"]
                except Exception as e:
                    st.error(f"Fehler beim Extrahieren der Spalten in Datei {uploaded_file.name}: {e}")
                    continue

                # Datumsspalte in deutschem Format (dd.mm.yyyy) formatieren
                try:
                    extracted_data["Datum"] = pd.to_datetime(extracted_data["Datum"], format="%d.%m.%Y", errors="coerce")
                except Exception as e:
                    st.error(f"Fehler bei der Umwandlung der Datumsspalte in Datei {uploaded_file.name}: {e}")
                    continue

                # Berechnung der Wertigkeiten
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

                # Monat und Jahr hinzufügen
                extracted_data["Monat"] = extracted_data["Datum"].dt.month
                extracted_data["Jahr"] = extracted_data["Datum"].dt.year

                # Daten zur Gesamtliste hinzufügen
                all_data = pd.concat([all_data, extracted_data], ignore_index=True)

            except Exception as e:
                st.error(f"Fehler beim Einlesen der Datei {uploaded_file.name}: {e}")

        # Export der Ergebnisse nach Monaten
        if not all_data.empty:
            output_file = "professionelle_touren_auswertung.xlsx"
            try:
                with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                    # Sortiere nach Jahr und Monat aufsteigend
                    sorted_data = all_data.sort_values(by=["Jahr", "Monat"])

                    for year, month in sorted_data[["Jahr", "Monat"]].drop_duplicates().values:
                        month_data = sorted_data[(sorted_data["Monat"] == month) & (sorted_data["Jahr"] == year)]
                        if not month_data.empty:
                            # Blattname als "Monat Jahr" in deutscher Sprache
                            month_name_german = {
                                "January": "Januar", "February": "Februar", "March": "März", "April": "April",
                                "May": "Mai", "June": "Juni", "July": "Juli", "August": "August",
                                "September": "September", "October": "Oktober", "November": "November", "December": "Dezember"
                            }
                            month_name = f"{month_name_german[calendar.month_name[month]]} {year}"

                            # Gruppieren nach Fahrer und detaillierte Darstellung
                            sheet_data = []
                            for (nachname, vorname), group in month_data.groupby(["Nachname", "Vorname"]):
                                # Fahrername hervorheben
                                sheet_data.append([f"{vorname} {nachname}", "", "", "", ""])
                                # Kopfzeile
                                sheet_data.append(["Datum", "Tour", "LKW2", "LKW3", "Verdienst"])
                                for _, row in group.iterrows():
                                    sheet_data.append([
                                        row["Datum"].strftime("%d.%m.%Y"),
                                        row["Tour"],
                                        row["LKW2"],
                                        row["LKW3"],
                                        row["Verdienst"]
                                    ])
                                # Gesamtverdienst hinzufügen
                                total_earnings = group["Verdienst"].sum()
                                sheet_data.append(["Gesamtverdienst", "", "", "", total_earnings])
                                sheet_data.append([])  # Leere Zeile als Trennung

                            # Erstellen eines DataFrames für das aktuelle Blatt
                            sheet_df = pd.DataFrame(sheet_data)

                            # Daten exportieren
                            sheet_df.to_excel(writer, index=False, sheet_name=month_name[:31])

                    # Automatische Spaltenbreite und Farbliche Anpassung
                    workbook = writer.book
                    light_gray = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
                    dark_gray = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
                    thin_border = Border(
                        left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin')
                    )

                    for sheet_name in workbook.sheetnames:
                        sheet = workbook[sheet_name]

                        # Setze Auto-Spaltenbreite
                        for col in sheet.columns:
                            max_length = max(len(str(cell.value)) for cell in col if cell.value)
                            col_letter = get_column_letter(col[0].column)
                            sheet.column_dimensions[col_letter].width = max_length + 2

                        # Farben und Umrandung
                        for row_idx, row in enumerate(sheet.iter_rows(), start=1):
                            if row_idx == 1:
                                # Fahrername fett und zentriert
                                for cell in row:
                                    cell.fill = dark_gray
                                    cell.font = Font(size=14, bold=True, color="FFFFFF")
                                    cell.alignment = Alignment(horizontal="center", vertical="center")
                                    cell.border = thin_border
                            elif row[0].value and "Datum" in str(row[0].value):
                                # Kopfzeile
                                for cell in row:
                                    cell.fill = light_gray
                                    cell.font = Font(bold=True)
                                    cell.border = thin_border
                            elif row[0].value and "Gesamtverdienst" in str(row[0].value):
                                # Gesamtverdienst
                                for cell in row:
                                    cell.fill = light_gray
                                    cell.font = Font(bold=True)
                                    cell.border = thin_border
                            else:
                                # Datenzeilen mit abwechselnden Farben
                                fill = light_gray if row_idx % 2 == 0 else None
                                for cell in row:
                                    if fill:
                                        cell.fill = fill
                                    cell.border = thin_border

                with open(output_file, "rb") as file:
                    st.download_button(
                        label="Download Auswertung",
                        data=file,
                        file_name="professionelle_touren_auswertung.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"Fehler beim Exportieren der Datei: {e}")

if __name__ == "__main__":
    main()
