import pandas as pd
import streamlit as st
import re
from datetime import datetime
from io import BytesIO

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

def format_sheet(writer, sheet_name, df):
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]

    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'center',
        'fg_color': '#D9E1F2',
        'border': 1
    })

    # Format Header
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)

    # Setze Spaltenbreite automatisch
    for i, col in enumerate(df.columns):
        max_length = max(df[col].astype(str).map(len).max(), len(str(col)))
        worksheet.set_column(i, i, max_length + 2)

def main():
    st.title("Zulage-Auswertung Sonderfahrzeuge")

    uploaded_files = st.file_uploader("Excel-Dateien hochladen", type=["xlsx"], accept_multiple_files=True)

    if uploaded_files:
        all_data = pd.DataFrame()

        for uploaded_file in uploaded_files:
            try:
                df = pd.read_excel(uploaded_file, sheet_name="Touren", header=None, skiprows=4)

                df = df.rename(columns={
                    3: "Nachname",
                    4: "Vorname",
                    11: "LKW",
                    13: "AZ-Kennung",
                    14: "Datum"
                })

                df = df[["Nachname", "Vorname", "LKW", "AZ-Kennung", "Datum"]]
                df = df[df["AZ-Kennung"].astype(str).str.contains("AZ", case=False, na=False)]
                df["Datum"] = pd.to_datetime(df["Datum"], errors="coerce")
                df = df[df["Datum"] >= pd.Timestamp("2025-01-01")]

                df["LKW"] = df["LKW"].apply(lambda x: f"E-{int(float(str(x).replace('E-', '')))}" if pd.notnull(x) else x)
                df["Verdienst"] = df.apply(calculate_earnings, axis=1)
                df["Monat"] = df["Datum"].dt.month
                df["Jahr"] = df["Datum"].dt.year

                all_data = pd.concat([all_data, df], ignore_index=True)

            except Exception as e:
                st.error(f"Fehler beim Verarbeiten von {uploaded_file.name}: {e}")

        if not all_data.empty:
            all_data.sort_values(by=["Nachname", "Vorname", "Jahr", "Monat"], inplace=True)
            st.dataframe(all_data)

            # 📦 Export vorbereiten
            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                # Blatt 1: Gesamtübersicht (alle Daten)
                all_data.to_excel(writer, sheet_name="Gesamtübersicht", index=False)
                format_sheet(writer, "Gesamtübersicht", all_data)

                # 📆 Monatsweise Blätter mit Gruppierung Fahrer + LKW
                grouped = all_data.groupby(["Jahr", "Monat"])

                for (jahr, monat), group in grouped:
                    sheet_name = f"{jahr}_{str(monat).zfill(2)}"

                    # Gruppieren: Fahrer + Fahrzeug
                    summary = group.groupby(["Nachname", "Vorname", "LKW"])["Verdienst"].sum().reset_index()
                    summary = summary.rename(columns={"Verdienst": "Monatssumme (€)"})

                    summary.to_excel(writer, sheet_name=sheet_name, index=False)
                    format_sheet(writer, sheet_name, summary)

            st.success("Excel mit Monatsblättern und Gruppierung erfolgreich erstellt.")

            st.download_button(
                label="📥 Excel-Datei herunterladen",
                data=output.getvalue(),
                file_name="Zulagen_Monatsauswertung.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
