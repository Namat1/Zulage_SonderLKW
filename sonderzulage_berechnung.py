import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

# Deutsche Monatsnamen
GERMAN_MONTHS = [
    "Dummy", "Januar", "Februar", "März", "April", "Mai", "Juni",
    "Juli", "August", "September", "Oktober", "November", "Dezember"
]

def get_german_month_name(month_number:int) -> str:
    return GERMAN_MONTHS[month_number]

# Fahrzeugart
def define_art(value):
    if value in [602, 156]:
        return "Gigaliner"
    elif value in [350, 620]:
        return "Tandem"
    elif value in [520, 266]:
        return "Gliederzug"
    return "Unbekannt"

# -------------------------------
# Personalnummer-Zuordnung
# -------------------------------
name_to_personalnummer = {
    "Adler": {"Philipp": "00041450"},
    "Auer": {"Frank": "00020795"},
    "Batkowski": {"Tilo": "00046601"},
    "Benabbes": {"Badr": "00048980"},
    "Biebow": {"Thomas": "00042004"},
    "Bläsing": {"Elmar": "00049093"},
    "Bursian": {"Ronny": "00025714"},
    "Buth": {"Sven": "00046673"},
    "Böhnke": {"Marcel": "00020833"},
    "Carstensen": {"Martin": "00042412"},
    "Chege": {"Moses Gichuru": "00046106"},
    "Dammasch": {"Bernd": "00019297"},
    "Demuth": {"Harry": "00020796"},
    "Doroszkiewicz": {"Bogumil": "00049132"},
    "Dürr": {"Holger": "00039164"},
    "Effenberger": {"Sven": "00030807"},
    "Engel": {"Raymond": "00033429"},
    "Fechner": {"Danny": "00043696", "Klaus": "00038278"},
    "Findeklee": {"Bernd": "00020804"},
    "Flint": {"Henryk": "00042414"},
    "Fuhlbrügge": {"Justin": "00046289"},
    "Gehrmann": {"Rayk": "00046702"},
    "Gheonea": {"Costel-Daniel": "00050877"},
    "Glanz": {"Björn": "00041914"},
    "Gnech": {"Torsten": "00018613"},
    "Greve": {"Nicole": "00040760"},
    "Guthmann": {"Fred": "00018328"},
    "Hagen": {"Andy": "00020271"},
    "Hartig": {"Sebastian": "00044120"},
    "Haus": {"David": "00046101"},
    "Heeser": {"Bernd": "00041916"},
    "Helm": {"Philipp": "00046685"},
    "Henkel": {"Bastian": "00048187"},
    "Holtz": {"Torsten": "00021159"},
    "Hirdina": {"Christopher": "00053400"},
    "Janikiewicz": {"Radoslaw": "00042159"},
    "Kelling": {"Jonas Ole": "00044140"},
    "Kleiber": {"Lutz": "00026255"},
    "Klemkow": {"Ralf": "00040634"},
    "Kollmann": {"Steffen": "00040988"},
    "König": {"Heiko": "00036341"},
    "Krazewski": {"Cezary": "00039463"},
    "Krieger": {"Christian": "00049092"},
    "Krull": {"Benjamin": "00044192"},
    "Lange": {"Michael": "00035407"},
    "Lewandowski": {"Kamil": "00041044"},
    "Likoonski": {"Vladimir": "00044766"},
    "Linke": {"Erich": "00048377"},
    "Lefkih": {"Houssni": "00052293"},
    "Ludolf": {"Michel": "00048814"},
    "Marouni": {"Ayyoub": "00048986"},
    "Mintel": {"Mario": "00046686"},
    "Ohlenroth": {"Nadja": "00042114"},
    "Ohms": {"Torsten": "00019300"},
    "Okoth": {"Tedy Omondi": "00046107"},
    "Oszmian": {"Jacub": "00039464"},
    "Paul": {"Toralf": "00010490"},
    "Pabst": {"Torsten": "00021976"},
    "Pawlak": {"Bartosz": "00036381"},
    "Piepke": {"Torsten": "00021390"},
    "Plinke": {"Killian": "00044137"},
    "Pogodski": {"Enrico": "00046668"},
    "Quint": {"Stefan": "00035718"},
    "Rimba": {"Rimba Gona": "00046108"},
    "Rheinschmitt": {"Ronald": "00053356"},
    "Sarwatka": {"Heiko": "00028747"},
    "Scheil": {"Eric-Rene": "00038579", "Rene": "00020851"},
    "Schlichting": {"Michael": "00021452"},
    "Schlutt": {"Hubert": "00020880", "Rene": "00042932"},
    "Schmieder": {"Steffen": "00046286"},
    "Schneider": {"Matthias": "00045495"},
    "Schulz": {"Julian": "00049130", "Stephan": "00041558"},
    "Singh": {"Jagtah": "00040902"},
    "Stoltz": {"Thorben": "00040991"},
    "Thal": {"Jannic": "00046006"},
    "Tumanow": {"Vasilli": "00045019"},
    "Wachnowski": {"Klaus": "00026019"},
    "Wendel": {"Danilo": "00048994"},
    "Wille": {"Rene": "00021393"},
    "Wisniewski": {"Krzysztof": "00046550"},
    "Zander": {"Jan": "00042454"},
    "Zosel": {"Ingo": "00026303"},
}

# --- Namens-Normalisierung & robuster Lookup ---
def _norm(s: str) -> str:
    return (s or "").strip().lower().replace("  ", " ")

def _norm_simple(s: str) -> str:
    return _norm(s).replace("-", " ").replace("  ", " ")

def get_personalnummer(nachname: str, vorname: str) -> str:
    """Robuste Zuordnung:
       - trim, casefold
       - toleriert Bindestriche/Mehrfachvornamen
       - versucht exakten Treffer, dann Startswith/Substring
       - fallback: erster Vorname (bis erstes Leerzeichen)
    """
    n_key = _norm_simple(nachname)
    v_key = _norm_simple(vorname)

    # 1) exakter Nachname-Schlüssel (case-insensitive)
    hit = None
    for ln, inner in name_to_personalnummer.items():
        if _norm_simple(ln) == n_key:
            # a) exakter Vorname
            for fn, pn in inner.items():
                if _norm_simple(fn) == v_key:
                    return pn
            # b) Vorname beginnt mit/enthält
            for fn, pn in inner.items():
                f_norm = _norm_simple(fn)
                if v_key.startswith(f_norm) or f_norm.startswith(v_key) or (f_norm in v_key) or (v_key in f_norm):
                    return pn
            # c) nur erster Vorname probieren
            if " " in v_key:
                first = v_key.split(" ", 1)[0]
                for fn, pn in inner.items():
                    if _norm_simple(fn).startswith(first):
                        return pn
            hit = "Unbekannt"
            break
    return hit or "Unbekannt"

# -------------------------------
# Styling (VERBESSERT)
# -------------------------------
def apply_styles(sheet):
    """Verbesserte Optik mit modernem Design"""
    
    # Moderne Farbpalette
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")  # Dunkelblau
    name_header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")  # Mittelblau
    subheader_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")  # Hellblau
    data_fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")  # Weiß
    data_fill_light = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")  # Sehr helles Grau
    total_fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")  # Grün für Summen
    
    # Rahmen
    thin_border = Border(
        left=Side(style='thin', color='CCCCCC'),
        right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'),
        bottom=Side(style='thin', color='CCCCCC')
    )
    
    medium_border = Border(
        left=Side(style='medium', color='1F4E78'),
        right=Side(style='medium', color='1F4E78'),
        top=Side(style='medium', color='1F4E78'),
        bottom=Side(style='medium', color='1F4E78')
    )
    
    is_first_in_block = True
    alternate_row = False

    for row_idx, row in enumerate(sheet.iter_rows(min_col=1, max_col=5), start=1):
        first_cell_value = str(row[0].value).strip() if row[0].value else ""

        # Gesamtverdienst-Zeile (Summe pro Person)
        if "Gesamtverdienst" in first_cell_value:
            for cell in row:
                cell.fill = total_fill
                cell.font = Font(bold=True, size=11, color="FFFFFF")
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.border = medium_border
                if cell.column == 5:
                    cell.number_format = '#,##0.00 €'
            is_first_in_block = True
            alternate_row = False

        # Erste Zeile eines neuen Mitarbeiter-Blocks (Name)
        elif is_first_in_block and first_cell_value:
            for cell in row:
                cell.fill = name_header_fill
                cell.font = Font(bold=True, size=13, color="FFFFFF")
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = medium_border
            is_first_in_block = False
            alternate_row = False

        # Spaltenüberschriften (Datum, Tour, LKW, Art, Verdienst)
        elif first_cell_value and not any(char.isdigit() for char in first_cell_value[:10]):
            # Prüfen ob es eine Header-Zeile ist (enthält Worte wie "Datum", "Tour", etc.)
            if any(keyword in first_cell_value for keyword in ["Datum", "Tour", "LKW", "Art", "Verdienst"]):
                for cell in row:
                    cell.fill = subheader_fill
                    cell.font = Font(bold=True, size=10, color="1F4E78")
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = thin_border
                alternate_row = False
            else:
                # Normale Namenszeile
                for cell in row:
                    cell.fill = name_header_fill
                    cell.font = Font(bold=True, size=11, color="FFFFFF")
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                    cell.border = thin_border

        # Datenzeilen (alternierende Farben)
        else:
            fill_color = data_fill_white if alternate_row else data_fill_light
            for cell in row:
                cell.fill = fill_color
                cell.font = Font(size=10, color="2C3E50")
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.border = thin_border
                
                # Währungsformat für Verdienst-Spalte
                if cell.column == 5:
                    try:
                        cell.value = float(cell.value)
                        cell.number_format = '#,##0.00 €'
                        # Grüne Schrift für positive Beträge
                        if cell.value > 0:
                            cell.font = Font(size=10, color="70AD47", bold=True)
                    except (ValueError, TypeError):
                        pass
            
            alternate_row = not alternate_row
            
            if not first_cell_value:
                is_first_in_block = True
                alternate_row = False

    # Spaltenbreiten optimieren
    for col in sheet.columns:
        max_length = max(len(str(cell.value) or "") for cell in col)
        col_letter = get_column_letter(col[0].column)
        adjusted_width = min(max_length + 4, 60)
        sheet.column_dimensions[col_letter].width = adjusted_width

    # Zeilenhöhe für bessere Lesbarkeit
    for row in range(1, sheet.max_row + 1):
        sheet.row_dimensions[row].height = 20

    # Erste Zeile ausblenden
    sheet.row_dimensions[1].hidden = True
    
    # Freeze Panes für bessere Navigation
    sheet.freeze_panes = "A3"

def format_date_with_german_weekday(date: pd.Timestamp) -> str:
    wochentage_mapping = {
        "Monday": "Montag", "Tuesday": "Dienstag", "Wednesday": "Mittwoch",
        "Thursday": "Donnerstag", "Friday": "Freitag", "Saturday": "Samstag",
        "Sunday": "Sonntag"
    }
    english_weekday = date.strftime("%A")
    german_weekday = wochentage_mapping.get(english_weekday, english_weekday)
    original_kw = int(date.strftime("%W"))
    adjusted_kw = original_kw + 1 if original_kw < 53 else 1
    return date.strftime(f"%d.%m.%Y ({german_weekday}, KW{adjusted_kw})")

def add_summary(sheet, summary_data, start_col=9, month_name=""):
    """Verbesserte Zusammenfassungstabelle mit modernem Design"""
    
    # Moderne Farbpalette für Summary
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")  # Dunkelblau
    subheader_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")  # Mittelblau
    cell_fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    cell_fill_light = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
    total_fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")  # Grün
    
    # Rahmen
    thin_border = Border(
        left=Side(style='thin', color='CCCCCC'),
        right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'),
        bottom=Side(style='thin', color='CCCCCC')
    )
    
    medium_border = Border(
        left=Side(style='medium', color='1F4E78'),
        right=Side(style='medium', color='1F4E78'),
        top=Side(style='medium', color='1F4E78'),
        bottom=Side(style='medium', color='1F4E78')
    )

    # Monatskopf über 3 Spalten
    c1 = sheet.cell(row=2, column=start_col, value="Auszahlung Monat:")
    c1.font = Font(bold=True, size=14, color="FFFFFF")
    c1.alignment = Alignment(horizontal="right", vertical="center")
    c1.fill = header_fill
    c1.border = medium_border

    c2 = sheet.cell(row=2, column=start_col+1, value=month_name or "Unbekannt")
    c2.font = Font(bold=True, size=14, color="FFFFFF")
    c2.alignment = Alignment(horizontal="center", vertical="center")
    c2.fill = header_fill
    c2.border = medium_border

    c3 = sheet.cell(row=2, column=start_col+2, value="")
    c3.font = Font(bold=True, size=14, color="FFFFFF")
    c3.alignment = Alignment(horizontal="left", vertical="center")
    c3.fill = header_fill
    c3.border = medium_border

    # Spaltenüberschriften
    headers = ["Name", "Personalnummer", "Gesamtverdienst (€)"]
    for i, h in enumerate(headers, start=start_col):
        cell = sheet.cell(row=3, column=i, value=h)
        cell.font = Font(bold=True, size=11, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = subheader_fill
        cell.border = medium_border

    # Sortierung nach Verdienst
    summary_data.sort(key=lambda x: x[2], reverse=True)

    # Datenzeilen mit alternierender Farbe
    for r, (name, personalnummer, total) in enumerate(summary_data, start=4):
        fill_color = cell_fill_white if r % 2 == 0 else cell_fill_light
        
        # Name
        nc = sheet.cell(row=r, column=start_col, value=name)
        nc.font = Font(bold=True, size=11, color="2C3E50")
        nc.alignment = Alignment(horizontal="left", vertical="center")
        nc.fill = fill_color
        nc.border = thin_border

        # Personalnummer
        pn_cell = sheet.cell(row=r, column=start_col+1)
        if personalnummer != "Unbekannt":
            pn_cell.value = int(personalnummer)
            pn_cell.number_format = '00000000'
        else:
            pn_cell.value = personalnummer
        pn_cell.font = Font(size=10, color="5A6C7D")
        pn_cell.alignment = Alignment(horizontal="center", vertical="center")
        pn_cell.fill = fill_color
        pn_cell.border = thin_border

        # Verdienst
        tc = sheet.cell(row=r, column=start_col+2, value=float(total))
        tc.font = Font(bold=True, size=11, color="70AD47")
        tc.fill = fill_color
        tc.border = thin_border
        tc.number_format = '#,##0.00 €'
        tc.alignment = Alignment(horizontal="right", vertical="center")

    # Gesamtsumme
    total_row = len(summary_data) + 4
    
    lab = sheet.cell(row=total_row, column=start_col, value="GESAMTSUMME")
    lab.font = Font(bold=True, size=12, color="FFFFFF")
    lab.alignment = Alignment(horizontal="right", vertical="center")
    lab.fill = total_fill
    lab.border = medium_border
    
    empty_cell = sheet.cell(row=total_row, column=start_col+1, value="")
    empty_cell.fill = total_fill
    empty_cell.border = medium_border
    
    sumcell = sheet.cell(row=total_row, column=start_col+2, value=sum(x[2] for x in summary_data))
    sumcell.number_format = '#,##0.00 €'
    sumcell.font = Font(bold=True, size=12, color="FFFFFF")
    sumcell.fill = total_fill
    sumcell.border = medium_border
    sumcell.alignment = Alignment(horizontal="right", vertical="center")

    # Zeilenhöhe für Summary
    for row in range(2, total_row + 1):
        sheet.row_dimensions[row].height = 22

    # Spaltenbreiten für Summary
    sheet.column_dimensions[get_column_letter(start_col)].width = 25
    sheet.column_dimensions[get_column_letter(start_col+1)].width = 18
    sheet.column_dimensions[get_column_letter(start_col+2)].width = 20

# -------------------------------
# App (UNVERÄNDERT)
# -------------------------------
def main():
    st.title("Zulage - Sonderfahrzeuge - Ab 2025")

    uploaded_files = st.file_uploader(
        "Lade eine oder mehrere Excel-Dateien hoch",
        type=["xlsx", "xls"],
        accept_multiple_files=True
    )

    if uploaded_files:
        all_data = pd.DataFrame()

        for uploaded_file in uploaded_files:
            try:
                df = pd.read_excel(uploaded_file, sheet_name="Touren", header=0)

                # AZ-Zeilen filtern (Spalte 14 -> Index 13)
                mask = df.iloc[:, 13].astype(str).str.contains(r'(?i)\bAZ\b', na=False)
                filtered_df = df[mask].copy()

                if not filtered_df.empty:
                    filtered_df["Datum"] = pd.to_datetime(filtered_df.iloc[:, 14], format="%d.%m.%Y", errors="coerce")
                    filtered_df = filtered_df[filtered_df["Datum"] >= pd.Timestamp("2025-01-01")]

                if filtered_df.empty:
                    st.warning(f"Keine passenden Daten in der Datei {uploaded_file.name} gefunden.")
                    continue

                # Relevante Spalten
                columns_to_extract = [0, 3, 4, 10, 11, 12, 14]
                extracted = filtered_df.iloc[:, columns_to_extract].copy()
                extracted.columns = ["Tour", "Nachname", "Vorname", "LKW1", "LKW", "Art", "Datum"]

                # Strings trimmen
                extracted["Nachname"] = extracted["Nachname"].astype(str).str.strip()
                extracted["Vorname"] = extracted["Vorname"].astype(str).str.strip()

                # LKW normalisieren + Art bestimmen
                extracted["LKW"] = extracted["LKW"].apply(lambda x: f"E-{x}" if pd.notnull(x) else x)
                extracted["Art"] = extracted["LKW"].apply(
                    lambda x: define_art(int(x.split("-")[1])) if pd.notnull(x) and "-" in str(x) and str(x).split("-")[1].isdigit() else "Unbekannt"
                )

                extracted["Datum"] = pd.to_datetime(extracted["Datum"], format="%d.%m.%Y", errors="coerce")

                # Tour ggf. aus Spalte Q (Index 16)
                if "Tour" in extracted.columns and 16 in filtered_df.columns:
                    extracted["Tour"] = extracted["Tour"].fillna(filtered_df.iloc[:, 16])

                # Verdienst
                def calculate_earnings(row):
                    earnings = 0
                    candidates = []
                    if pd.notnull(row["LKW1"]) and str(row["LKW1"]).isdigit():
                        candidates.append(int(row["LKW1"]))
                    if pd.notnull(row["LKW"]) and "-" in str(row["LKW"]):
                        tail = str(row["LKW"]).split("-")[1]
                        if tail.isdigit():
                            candidates.append(int(tail))
                    for v in candidates:
                        if v in [602, 156]:
                            earnings += 40
                        elif v in [620, 350, 520, 266]:
                            earnings += 20
                    return earnings

                extracted["Verdienst"] = extracted.apply(calculate_earnings, axis=1)
                extracted["Monat"] = extracted["Datum"].dt.month
                extracted["Jahr"] = extracted["Datum"].dt.year

                all_data = pd.concat([all_data, extracted], ignore_index=True)

            except Exception as e:
                st.error(f"Fehler beim Einlesen der Datei {uploaded_file.name}: {e}")

        if not all_data.empty:
            output_file = "touren_auswertung_korrekt.xlsx"
            try:
                with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                    sorted_data = all_data.sort_values(by=["Jahr", "Monat", "Nachname", "Vorname"])

                    for year, month in sorted_data[["Jahr", "Monat"]].drop_duplicates().values:
                        month_data = sorted_data[(sorted_data["Jahr"] == year) & (sorted_data["Monat"] == month)].copy()
                        if month_data.empty:
                            continue

                        sheet_name = f"{get_german_month_name(month)} {year}"
                        sheet_data = []
                        summary_data = []

                        # Gruppieren pro Person
                        for (nachname, vorname), group in month_data.groupby(["Nachname", "Vorname"], dropna=False):
                            vn = (vorname or "").strip()
                            nn = (nachname or "").strip()

                            total_earnings = float(group["Verdienst"].sum())
                            personalnummer = get_personalnummer(nn, vn)

                            summary_data.append([f"{vn} {nn}".strip(), personalnummer, total_earnings])

                            # Block-Kopf
                            sheet_data.append([f"{vn} {nn}".strip(), "", "", "", ""])
                            sheet_data.append(["Datum", "Tour", "LKW", "Art", "Verdienst"])

                            for _, row in group.iterrows():
                                dt = pd.to_datetime(row["Datum"]) if pd.notnull(row["Datum"]) else pd.NaT
                                formatted_date = format_date_with_german_weekday(dt) if pd.notnull(dt) else ""
                                sheet_data.append([
                                    formatted_date,
                                    row["Tour"],
                                    row["LKW"],
                                    row["Art"],
                                    float(row["Verdienst"])
                                ])

                            sheet_data.append(["Gesamtverdienst", "", "", "", total_earnings])
                            sheet_data.append([])

                        # Blatt schreiben
                        pd.DataFrame(sheet_data).to_excel(writer, index=False, sheet_name=sheet_name[:31])
                        sheet = writer.sheets[sheet_name[:31]]

                        # Zusammenfassung & Styling
                        add_summary(sheet, summary_data, start_col=9, month_name=sheet_name)
                        apply_styles(sheet)

                with open(output_file, "rb") as fh:
                    st.download_button(
                        label="Download Auswertung",
                        data=fh,
                        file_name="Zulage_Sonderfahrzeuge_2025.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"Fehler beim Exportieren der Datei: {e}")

if __name__ == "__main__":
    main()
