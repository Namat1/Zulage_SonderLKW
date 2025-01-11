
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
    # ... (gekürzt für Lesbarkeit)
}

def apply_styles(sheet):
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    for col in sheet.columns:
        max_length = max(len(str(cell.value) or "") for cell in col)
        col_letter = get_column_letter(col[0].column)
        sheet.column_dimensions[col_letter].width = max_length + 2

def format_date_with_weekday_and_kw(date):
    if pd.isnull(date):
        return ""
    weekday = date.strftime("%A")
    iso_calendar = date.isocalendar()
    kw = iso_calendar.week
    return f"{date.strftime('%d.%m.%Y')} ({weekday}, KW{kw})"

def main():
    st.title("Zulage - Sonderfahrzeuge - Ab 2025")
    uploaded_files = st.file_uploader("Lade Excel-Dateien hoch", type=["xlsx"], accept_multiple_files=True)

    if uploaded_files:
        all_data = pd.DataFrame()
        for uploaded_file in uploaded_files:
            try:
                # Excel-Datei laden
                df = pd.read_excel(uploaded_file, sheet_name="Touren")
                
                # Prüfen, ob Spalte 15 vorhanden ist
                if df.shape[1] >= 15:
                    df["Datum"] = pd.to_datetime(df.iloc[:, 14], errors="coerce")
                    df["Datum_Formatted"] = df["Datum"].apply(format_date_with_weekday_and_kw)
                else:
                    st.error(f"Die Datei {uploaded_file.name} hat keine ausreichenden Spalten.")
                    continue

                all_data = pd.concat([all_data, df], ignore_index=True)
            except Exception as e:
                st.error(f"Fehler beim Einlesen der Datei {uploaded_file.name}: {e}")

        if not all_data.empty:
            output_file = "Zulage_Auswertung.xlsx"
            with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                all_data.to_excel(writer, index=False, sheet_name="Daten")
                sheet = writer.sheets["Daten"]
                apply_styles(sheet)

            with open(output_file, "rb") as file:
                st.download_button("Download Excel-Datei", file, file_name=output_file)

if __name__ == "__main__":
    main()
