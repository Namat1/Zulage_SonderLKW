
import streamlit as st
import pandas as pd
import io

st.title("Füngers-Zulagen pro Monat & Fahrer")

uploaded_files = st.file_uploader("Excel-Dateien hochladen", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    alle_eintraege = []

    for file in uploaded_files:
        try:
            df = pd.read_excel(file, sheet_name="Touren", header=None)
            df = df.iloc[4:]  # ab Zeile 5
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
                    alle_eintraege.append({
                        "Nachname": name,
                        "Vorname": vorname,
                        "Datum": datum.date(),
                        "Kommentar": kommentar,
                        "Verdienst": 20,
                        "Monat": datum.strftime("%B %Y")
                    })

        except Exception as e:
            st.error(f"Fehler beim Verarbeiten der Datei {file.name}: {e}")

    if alle_eintraege:
        df_gesamt = pd.DataFrame(alle_eintraege)
        st.success(f"{len(df_gesamt)} gültige Füngers-Zulagen erkannt.")
        st.dataframe(df_gesamt)

        # Schreibe pro Monat ein Blatt
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            for monat, monat_df in df_gesamt.groupby("Monat"):
                sheet_df = []
                for (nach, vor), gruppe in monat_df.groupby(["Nachname", "Vorname"]):
                    sheet_df.append([f"{vor} {nach}", "", "", "", ""])
                    sheet_df.append(["Datum", "Kommentar", "Verdienst", "", ""])
                    for _, r in gruppe.iterrows():
                        sheet_df.append([r["Datum"], r["Kommentar"], r["Verdienst"], "", ""])
                    sheet_df.append(["Gesamt", "", gruppe["Verdienst"].sum(), "", ""])
                    sheet_df.append([])

                pd.DataFrame(sheet_df).to_excel(writer, index=False, sheet_name=monat[:31])

        st.download_button("Excel mit Monatsblättern herunterladen", output.getvalue(), file_name="füngers_monate.xlsx")

    else:
        st.warning("Keine gültigen Füngers-Zulagen gefunden.")
