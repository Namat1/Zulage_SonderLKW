import pandas as pd
import streamlit as st

def main():
    st.title("Touren-Auswertung")
    
    # Hochladen der Excel-Datei
    uploaded_file = st.file_uploader("Lade eine Excel-Datei hoch", type=["xlsx", "xls"])
    
    if uploaded_file:
        try:
            # Einlesen der Excel-Tabelle
            df = pd.read_excel(uploaded_file, sheet_name="Touren", header=0)
            st.write("Daten erfolgreich geladen. Vorschau:")
            st.dataframe(df.head())
        except Exception as e:
            st.error(f"Fehler beim Einlesen der Datei: {e}")
            return
        
        # Filter auf Spalte 14 (Werte Az, AZ, az)
        try:
            filtered_df = df[df.iloc[:, 13].str.contains(r'(?i)\b(AZ)\b', na=False)]
            if filtered_df.empty:
                st.warning("Keine passenden Daten gefunden!")
                return
        except Exception as e:
            st.error(f"Fehler beim Filtern der Daten: {e}")
            return
        
        # Relevante Spalten extrahieren (ohne Spalte 13)
        columns_to_extract = [0, 3, 4, 10, 11, 12, 14]  # Spalten: 1, 4, 5, 11, 12, 14, 15
        try:
            extracted_data = filtered_df.iloc[:, columns_to_extract]
            extracted_data.columns = ["Tour", "Nachname", "Vorname", "LKW1", "LKW2", "LKW3", "Datum"]
        except Exception as e:
            st.error(f"Fehler beim Extrahieren der Spalten: {e}")
            return
        
        # Datumsspalte in ein Datumsformat umwandeln
        try:
            extracted_data["Datum"] = pd.to_datetime(extracted_data["Datum"], format="%d.%m.%Y", errors="coerce")
        except Exception as e:
            st.error(f"Fehler bei der Umwandlung der Datumsspalte: {e}")
            return
        
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
        
        # Gruppierung nach Fahrer und Monat
        try:
            extracted_data["Monat"] = extracted_data["Datum"].dt.to_period("M")
            summary = extracted_data.groupby(["Nachname", "Vorname", "Monat"]).agg(
                {"Verdienst": "sum"}
            ).reset_index()
        except Exception as e:
            st.error(f"Fehler bei der Berechnung der Zusammenfassung: {e}")
            return
        
        # Daten anzeigen
        st.subheader("Gefilterte Daten")
        st.dataframe(extracted_data)
        
        st.subheader("Monatliche Verdienste")
        st.dataframe(summary)
        
        # Download der neuen Excel
        output_file = "auswertung.xlsx"
        try:
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
        except Exception as e:
            st.error(f"Fehler beim Exportieren der Datei: {e}")
            return

if __name__ == "__main__":
    main()
