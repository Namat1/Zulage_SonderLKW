
import streamlit as st
import pandas as pd
import io

st.title("Füngers-Zulagen-Auswertung")

uploaded_files = st.file_uploader("Excel-Dateien hochladen", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    zulagen_liste = []

    for file in uploaded_files:
        try:
            df = pd.read_excel(file, sheet_name="Touren", header=None)
            df = df.iloc[4:]  # ab Zeile 5
            df.columns = range(df.shape[1])  # sichere Spaltenindexierung

            for _, row in df.iterrows():
                kommentar = str(row[15]) if 15 in row and pd.notnull(row[15]) else ""
                name = row[3] if 3 in row else None
                vorname = row[4] if 4 in row else None

                if (
                    "füngers" in kommentar.lower()
                    and pd.notnull(name)
                    and pd.notnull(vorname)
                ):
                    zulagen_liste.append({
                        "Nachname": name,
                        "Vorname": vorname,
                        "Kommentar": kommentar,
                        "Verdienst": 20
                    })

        except Exception as e:
            st.error(f"Fehler beim Verarbeiten der Datei {file.name}: {e}")

    if zulagen_liste:
        df_zulagen = pd.DataFrame(zulagen_liste)
        st.success(f"{len(df_zulagen)} gültige Füngers-Zulagen gefunden.")
        st.dataframe(df_zulagen)

        # Excel-Download ermöglichen
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_zulagen.to_excel(writer, index=False, sheet_name="Füngers-Zulagen")
        st.download_button("Excel herunterladen", output.getvalue(), file_name="füngers_zulagen.xlsx")

    else:
        st.warning("Keine gültigen Füngers-Zulagen gefunden (Name oder Vorname fehlt?).")
