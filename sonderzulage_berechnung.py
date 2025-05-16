
import pandas as pd
import streamlit as st
import re
from datetime import datetime

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

                filtered_df = df[df["AZ-Kennung"].astype(str).str.contains("AZ", case=False, na=False)]
                filtered_df["Datum"] = pd.to_datetime(filtered_df["Datum"], errors="coerce")
                filtered_df = filtered_df[filtered_df["Datum"] >= pd.Timestamp("2025-01-01")]

                filtered_df["LKW"] = filtered_df["LKW"].apply(lambda x: f"E-{int(float(str(x).replace('E-', '')))}" if pd.notnull(x) else x)
                filtered_df["Verdienst"] = filtered_df.apply(calculate_earnings, axis=1)

                all_data = pd.concat([all_data, filtered_df], ignore_index=True)

            except Exception as e:
                st.error(f"Fehler beim Verarbeiten von {uploaded_file.name}: {e}")

        if not all_data.empty:
            st.dataframe(all_data)
            st.success("Daten erfolgreich verarbeitet.")

if __name__ == "__main__":
    main()
