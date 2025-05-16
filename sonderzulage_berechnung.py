
import pandas as pd
import streamlit as st
from datetime import datetime

def calculate_earnings(row):
    try:
        if pd.notnull(row["LKW"]) and "-" in str(row["LKW"]):
            nummer = int(str(row["LKW"]).split("-")[1])
            if nummer in [266, 520, 620, 350]:
                return 20
            elif nummer in [602, 156]:
                return 40
    except:
        pass
    return 0

def main():
    st.title("Zulage-Auswertung Sonderfahrzeuge")

    uploaded_files = st.file_uploader("Excel-Dateien hochladen", type=["xlsx"], accept_multiple_files=True)

    if uploaded_files:
        all_data = pd.DataFrame()

        for uploaded_file in uploaded_files:
            try:
                # ab Zeile 5 einlesen (index 4), keine Header-Zeile nutzen
                df = pd.read_excel(uploaded_file, sheet_name="Touren", header=None, skiprows=4)

                # Spalten zuweisen
                df = df.rename(columns={
                    3: "Nachname",
                    4: "Vorname",
                    11: "LKW",
                    13: "AZ-Kennung",
                    14: "Datum"
                })

                df = df[["Nachname", "Vorname", "LKW", "AZ-Kennung", "Datum"]]

                # Filter auf AZ
                filtered_df = df[df["AZ-Kennung"].astype(str).str.contains("AZ", case=False, na=False)]
                filtered_df["Datum"] = pd.to_datetime(filtered_df["Datum"], errors="coerce")
                filtered_df = filtered_df[filtered_df["Datum"] >= pd.Timestamp("2025-01-01")]

                # LKW-Format und Verdienst
                filtered_df["LKW"] = filtered_df["LKW"].apply(lambda x: f"E-{x}" if pd.notnull(x) else x)
                filtered_df["Verdienst"] = filtered_df.apply(calculate_earnings, axis=1)

                all_data = pd.concat([all_data, filtered_df], ignore_index=True)

            except Exception as e:
                st.error(f"Fehler beim Verarbeiten von {uploaded_file.name}: {e}")

        if not all_data.empty:
            st.dataframe(all_data)
            st.success("Daten erfolgreich verarbeitet.")

if __name__ == "__main__":
    main()
