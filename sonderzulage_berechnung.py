
import streamlit as st
import pandas as pd
import io
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

st.title("Füngers-Zulagen Auswertung – Monatsübersicht")

uploaded_files = st.file_uploader("Excel-Dateien hochladen", type=["xlsx"], accept_multiple_files=True)

# Deutsche Monatsnamen (mit Nummer für Sortierung)
german_months = {
    1: "Januar", 2: "Februar", 3: "März", 4: "April",
    5: "Mai", 6: "Juni", 7: "Juli", 8: "August",
    9: "September", 10: "Oktober", 11: "November", 12: "Dezember"
}

if uploaded_files:
    eintraege = []

    for file in uploaded_files:
        try:
            df = pd.read_excel(file, sheet_name="Touren", header=None)
            df = df.iloc[4:]
            df.columns = range(df.shape[1])

            for _, row in df.iterrows():
                kommentar = str(row[15]) if 15 in row and pd.notnull(row[15]) else ""
                name = row[3] if 3 in row else None
                vorname = row[4] if 4 in row else None
                datum = pd.to_datetime(row[14], errors='coerce') if 14 in row else None

                if (
                    "füngers" in kommentar.lower()
                    and pd.notnull(name)
                    and pd.notnull(vorname)
                    and pd.notnull(datum)
                ):
                    monat_index = datum.month
                    jahr = datum.year
                    monat_name = german_months[monat_index]
                    eintraege.append({
                        "Nachname": name,
                        "Vorname": vorname,
                        "Datum": datum.strftime("%d.%m.%Y"),
                        "Kommentar": kommentar,
                        "Verdienst": 20,
                        "Monat": f"{monat_index:02d}-{jahr}_{monat_name} {jahr}"
                    })

        except Exception as e:
            st.error(f"Fehler in Datei {file.name}: {e}")

    if eintraege:
        df_gesamt = pd.DataFrame(eintraege)
        st.success(f"{len(df_gesamt)} gültige Füngers-Zulagen erkannt.")
        st.dataframe(df_gesamt)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for monat_key in sorted(df_gesamt["Monat"].unique()):
                df_monat = df_gesamt[df_gesamt["Monat"] == monat_key]
                zeilen = []
                for (nach, vor), gruppe in df_monat.groupby(["Nachname", "Vorname"]):
                    zeilen.append([f"{vor} {nach}"])
                    zeilen.append(["Datum", "Kommentar", "Verdienst"])
                    for _, r in gruppe.iterrows():
                        zeilen.append([r["Datum"], r["Kommentar"], r["Verdienst"]])
                    zeilen.append(["Gesamt", "", gruppe["Verdienst"].sum()])
                    zeilen.append([])

                df_sheet = pd.DataFrame(zeilen)
                sheet_name = monat_key.split("_")[1][:31]
                df_sheet.to_excel(writer, index=False, sheet_name=sheet_name)

                # Formatieren
                sheet = writer.sheets[sheet_name]
                thin = Border(left=Side(style='thin'), right=Side(style='thin'),
                              top=Side(style='thin'), bottom=Side(style='thin'))
                for row in sheet.iter_rows():
                    for cell in row:
                        cell.font = Font(name="Calibri", size=11)
                        cell.alignment = Alignment(horizontal="left", vertical="center")
                        cell.border = thin
                        if cell.row == 2:
                            cell.fill = PatternFill("solid", fgColor="95b3d7")
                            cell.font = Font(bold=True)
                        if "Gesamt" in str(cell.value):
                            cell.font = Font(bold=True)
                            cell.fill = PatternFill("solid", fgColor="c7b7b3")

                # Spaltenbreiten
                for col in sheet.columns:
                    max_length = max(len(str(cell.value) or "") for cell in col)
                    col_letter = get_column_letter(col[0].column)
                    sheet.column_dimensions[col_letter].width = max_length + 3

        st.download_button("Excel-Datei herunterladen", output.getvalue(), file_name="füngers_monatsauswertung_final.xlsx")

    else:
        st.warning("Keine gültigen Füngers-Zulagen gefunden.")
