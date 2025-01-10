import pandas as pd
import streamlit as st

def main():
    st.title("Touren-Auswertung")
    
    # Hochladen der Excel-Datei
    uploaded_file = st.file_uploader("Lade eine Excel-Datei hoch", type=["xlsx", "xls"])
    
    if uploaded_file:
        # Einlesen der Excel-Tabelle
        try:
            df = pd.read_excel(uploaded_file, sheet_name="Touren", header=0)
        except Exception as e:
            st.error(f"Fehler beim Einlesen der Datei: {e}")
            return
        
        # Filter auf Spalte 14 (Werte Az, AZ, az)
        filtered_df = df[df.iloc[:, 13].str.contains(r'(?i)\b(AZ)\b', na=False)]
        
        # Relevante Spalten extrahieren
        columns_to_extract = [0, 3, 4, 10, 11, 12, 13, 14]  # Spalten: 1, 4, 5, 11, 12, 13, 14, 15
        extracted_data = filtered_df.iloc[:, columns_to_extract]
        extracted_data.columns = ["Tour", "Nachname", "Vorname", "LKW1", "LKW2", "LKW3", "LKW4", "Datum"]
        
        # Berechnung der Wertigkeiten
        def calculate_earnings(row):
            lkw_values = [row["LKW1"], row["LKW2"], row["LKW3"], row["LKW4"]]
            earnings = 0
            for value in lkw_values:
                if value in [602, 156]:
                    earnings += 40
                elif value in [620, 350, 520]:
                    earnings += 20
            return earnings
        
        extracted_data["Verdienst"] = extracted_data.apply(calculate_earnings, axis=1)
        
        # Gruppierung nach Fahrer und Monat
        extracted_data["Monat"] = pd.to_datetime(extracted_data["Datum"]).dt.to_period("M")
        summary = extracted_data.groupby(["Nachname", "Vorname", "Monat"]).agg(
            {"Verdienst": "sum"}
        ).reset_index()
        
        # Daten anzeigen
        st.subheader("Gefilterte Daten")
        st.dataframe(extracted_data)
        
        st.subheader("Monatliche Verdienste")
        st.dataframe(summary)
        
        # Download der neuen Excel
        output_file = "auswertung.xlsx"
        with pd.ExcelWriter(output_file) as writer:
            extracted_data.to_excel(writer, index=False, sheet_name="Gefilterte Daten")
            summary.to_excel(writer, index=False, sheet_name="Zusammenfassung")
        
        with open(output_file, "rb") as file:
            st.download_button(
                label="Download Auswertung",
                data=file,
                file_name="auswertung.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
