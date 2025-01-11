import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


def restructure_data(all_data):
    """
    Restrukturiert die Daten, sodass Fahrer nebeneinander dargestellt werden.
    """
    # Alle Fahrer identifizieren
    unique_fahrers = all_data[["Nachname", "Vorname"]].drop_duplicates()
    unique_fahrers["Fahrer"] = unique_fahrers["Vorname"] + " " + unique_fahrers["Nachname"]

    # Pivot-Struktur erstellen
    result = pd.DataFrame()
    for _, (nachname, vorname, fahrer) in unique_fahrers.iterrows():
        fahrer_data = all_data[(all_data["Nachname"] == nachname) & (all_data["Vorname"] == vorname)]
        fahrer_data = fahrer_data[["Datum", "Tour", "LKW2", "LKW3", "Verdienst"]]
        fahrer_data.columns = [f"{fahrer}: Datum", f"{fahrer}: Tour", f"{fahrer}: LKW2", f"{fahrer}: LKW3", f"{fahrer}: Verdienst"]

        if result.empty:
            result = fahrer_data
        else:
            result = pd.concat([result.reset_index(drop=True), fahrer_data.reset_index(drop=True)], axis=1)

    return result


def export_to_excel(all_data):
    """
    Exportiert die restrukturierten Daten in eine Excel-Datei.
    """
    output_file = "fahrer_nebeneinander.xlsx"

    try:
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            # Daten restrukturieren
            restructured_data = restructure_data(all_data)
            restructured_data.to_excel(writer, index=False, sheet_name="Fahrer nebeneinander")

            # Stil anwenden
            workbook = writer.book
            sheet = workbook.active

            # Farben und Stil
            header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")  # Hellgrau für Kopfzeilen
            thin_border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )
            for row in sheet.iter_rows():
                for cell in row:
                    cell.border = thin_border
                    if cell.row == 1:  # Kopfzeile
                        cell.fill = header_fill
                        cell.font = Font(bold=True)

    except Exception as e:
        st.error(f"Fehler beim Exportieren der Datei: {e}")

    return output_file


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
            output_file = export_to_excel(all_data)
            with open(output_file, "rb") as file:
                st.download_button(
                    label="Download Auswertung",
                    data=file,
                    file_name="fahrer_nebeneinander.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )


if __name__ == "__main__":
    main()
