import pandas as pd
import streamlit as st

def main():
    st.title("Touren-Auswertung für mehrere Dateien")

    # Mehrere Dateien hochladen
    uploaded_files = st.file_uploader("Lade eine oder mehrere Excel-Dateien hoch", type=["xlsx", "xls"], accept_multiple_files=True)

    if uploaded_files:
        all_results = []  # Liste zur Speicherung der Ergebnisse für jede Datei

        for uploaded_file in uploaded_files:
            try:
                # Lesen des Blattes "Touren" aus der aktuellen Datei
                df = pd.read_excel(uploaded_file, sheet_name="Touren", header=0)
                st.write(f"Datei: {uploaded_file.name} - Blatt 'Touren' erfolgreich geladen.")

                # Filter auf Spalte 14 (Werte Az, AZ, az)
                try:
                    filtered_df = df[df.iloc[:, 13].str.contains(r'(?i)\b(AZ)\b', na=False)]
                    if filtered_df.empty:
                        st.warning(f"Keine passenden Daten im Blatt 'Touren' der Datei {uploaded_file.name} gefunden.")
                        continue
                except Exception as e:
                    st.error(f"Fehler beim Filtern der Daten in Datei {uploaded_file.name}: {e}")
                    continue

                # Relevante Spalten extrahieren
                columns_to_extract = [0, 3, 4, 10, 11, 12, 14]  # Spalten: 1, 4, 5, 11, 12, 14, 15
                try:
                    extracted_data = filtered_df.iloc[:, columns_to_extract]
                    extracted_data.columns = ["Tour", "Nachname", "Vorname", "LKW1", "LKW2", "LKW3", "Datum"]
                except Exception as e:
                    st.error(f"Fehler beim Extrahieren der Spalten in Datei {uploaded_file.name}: {e}")
                    continue

                # Datumsspalte in deutschem Format (dd.mm.yyyy) formatieren
                try:
                    extracted_data["Datum"] = pd.to_datetime(extracted_data["Datum"], format="%d.%m.%Y", errors="coerce")
                    extracted_data["Datum"] = extracted_data["Datum"].dt.strftime("%d.%m.%Y")
                except Exception as e:
                    st.error(f"Fehler bei der Umwandlung der Datumsspalte in Datei {uploaded_file.name}: {e}")
                    continue

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

                # Gruppierung nach Monat
                try:
                    extracted_data["Monat"] = pd.to_datetime(extracted_data["Datum"], format="%d.%m.%Y").dt.to_period("M")
                    summary = extracted_data.groupby(["Monat", "Nachname", "Vorname"]).agg(
                        {"Verdienst": "sum"}
                    ).reset_index()
                    summary["Monat"] = summary["Monat"].astype(str)  # Monat in lesbarem Format
                except Exception as e:
                    st.error(f"Fehler bei der Berechnung der Zusammenfassung in Datei {uploaded_file.name}: {e}")
                    continue

                # Speichere die Ergebnisse für späteren Export
                all_results.append({
                    "file": uploaded_file.name,
                    "data": summary
                })

            except Exception as e:
                st.error(f"Fehler beim Einlesen der Datei {uploaded_file.name}: {e}")

        # Export der Ergebnisse in eine Excel-Datei
        if all_results:
            output_file = "auswertung_touren.xlsx"
            try:
                with pd.ExcelWriter(output_file) as writer:
                    for result in all_results:
                        sheet_name = result["file"][:31]  # Begrenzung des Blattnamens auf 31 Zeichen
                        result["data"].to_excel(writer, index=False, sheet_name=sheet_name)

                with open(output_file, "rb") as file:
                    st.download_button(
                        label="Download Auswertung",
                        data=file,
                        file_name="auswertung_touren.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"Fehler beim Exportieren der Datei: {e}")

if __name__ == "__main__":
    main()
