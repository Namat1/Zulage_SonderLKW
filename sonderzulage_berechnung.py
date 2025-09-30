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

# Fahrzeugart (Berechnungslogik bleibt unverändert)
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

# --- robuster Namens-Lookup (ohne Einfluss auf die Berechnung) ---
def _norm(s: str) -> str:
    return (s or "").strip().lower().replace("  ", " ")

def _norm_simple(s: str) -> str:
    return _norm(s).replace("-", " ").replace("  ", " ")

def get_personalnummer(nachname: str, vorname: str) -> str:
    n_key = _norm_simple(nachname)
    v_key = _norm_simple(vorname)

    for ln, inner in name_to_personalnummer.items():
        if _norm_simple(ln) == n_key:
            # exakter Vorname
            for fn, pn in inner.items():
                if _norm_simple(fn) == v_key:
                    return pn
            # begins-with/contains
            for fn, pn in inner.items():
                f_norm = _norm_simple(fn)
                if v_key.startswith(f_norm) or f_norm.startswith(v_key) or (f_norm in v_key) or (v_key in f_norm):
                    return pn
            # nur erster Vorname probieren
            if " " in v_key:
                first = v_key.split(" ", 1)[0]
                for fn, pn in inner.items():
                    if _norm_simple(fn).startswith(first):
                        return pn
            return "Unbekannt"
    return "Unbekannt"

# -------------------------------
# Optik / Excel-Styling (nur Ausgabe, Logik bleibt)
# -------------------------------
def apply_styles(sheet):
    thin = Border(left=Side(style="thin"), right=Side(style="thin"),
                  top=Side(style="thin"), bottom=Side(style="thin"))

    # Farben
    c_name      = "95b3d7"  # Kopfzeile je Person (blau)
    c_head      = "e6e6e6"  # Spaltenüberschrift "Datum/Tour/..."
    c_row_a     = "ffffff"  # Datenzeile A
    c_row_b     = "f7f7f7"  # Datenzeile B (Zebra)
    c_total     = "ddd0cb"  # Gesamtverdienst-Zeile
    c_summary   = "f2ece8"  # Fallback

    fill_name     = PatternFill("solid", start_color=c_name,    end_color=c_name)
    fill_head     = PatternFill("solid", start_color=c_head,    end_color=c_head)
    fill_a        = PatternFill("solid", start_color=c_row_a,   end_color=c_row_a)
    fill_b        = PatternFill("solid", start_color=c_row_b,   end_color=c_row_b)
    fill_total    = PatternFill("solid", start_color=c_total,   end_color=c_total)
    fill_fallback = PatternFill("solid", start_color=c_summary, end_color=c_summary)

    in_block = False
    zebra_toggle = False

    # Wir formatieren nur die ersten 5 Spalten (Detailtabelle)
    for row in sheet.iter_rows(min_col=1, max_col=5):
        first = (str(row[0].value).strip() if row[0].value is not None else "")

        # 1) Name-Zeile (Blockstart)
        if first and first not in ("Datum", "Gesamtverdienst"):
            in_block = True
            zebra_toggle = False
            for cell in row:
                cell.fill = fill_name
                cell.font = Font(bold=True, size=12, italic=True)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin
            row[0].alignment = Alignment(horizontal="left", vertical="center")
            continue

        # 2) Spaltenüberschriften innerhalb des Blocks
        if in_block and first == "Datum":
            for cell in row:
                cell.fill = fill_head
                cell.font = Font(bold=True, size=11)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin
            zebra_toggle = False
            continue

        # 3) Gesamtverdienst – Blockende
        if in_block and first == "Gesamtverdienst":
            for idx, cell in enumerate(row, start=1):
                cell.fill = fill_total
                cell.font = Font(bold=True, size=11)
                cell.alignment = Alignment(horizontal=("left" if idx == 1 else "right"), vertical="center")
                cell.border = thin
                if idx == 5:
                    cell.number_format = '#,##0.00 €'
            in_block = False
            continue

        # 4) Datenzeilen im Block (Zebra)
        if in_block:
            zebra_toggle = not zebra_toggle
            current_fill = fill_a if zebra_toggle else fill_b
            for idx, cell in enumerate(row, start=1):
                cell.fill = current_fill
                cell.font = Font(size=11)
                if idx == 1:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                elif idx in (2, 3, 4):
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    try:
                        cell.value = float(cell.value)
                        cell.number_format = '#,##0.00 €'
                    except (TypeError, ValueError):
                        pass
                cell.border = thin
            continue

        # 5) Fallback (z. B. leere Zeilen)
        for cell in row:
            cell.fill = fill_fallback if first else PatternFill("solid", start_color="FFFFFF", end_color="FFFFFF")
            cell.font = Font(size=11)
            cell.alignment = Alignment(horizontal="left", vertical="center")
            cell.border = thin

    # Spaltenbreiten (Detailtabelle)
    widths = {1: 30, 2: 12, 3: 10, 4: 14, 5: 14}
    for col_idx, width in widths.items():
        sheet.column_dimensions[get_column_letter(col_idx)].width = width

    # Erste Zeile (DF-Header) ausblenden
    sheet.row_dimensions[1].hidden = True

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
    header_fill = PatternFill("solid", start_color="95b3d7", end_color="95b3d7")
    head2_fill  = PatternFill("solid", start_color="c7b7b3", end_color="c7b7b3")
    row_a_fill  = PatternFill("solid", start_color="ffffff", end_color="ffffff")
    row_b_fill  = PatternFill("solid", start_color="f7f7f7", end_color="f7f7f7")
    total_fill  = PatternFill("solid", start_color="ddd0cb", end_color="ddd0cb")
    thin = Border(left=Side(style="thin"), right=Side(style="thin"),
                  top=Side(style="thin"), bottom=Side(style="thin"))

    # Kopf "Auszahlung Monat"
    r = 2
    c1 = sheet.cell(row=r, column=start_col,     value="Auszahlung Monat:")
    c2 = sheet.cell(row=r, column=start_col + 1, value=month_name or "Unbekannt")
    c3 = sheet.cell(row=r, column=start_col + 2, value="")
    for c in (c1, c2, c3):
        c.font = Font(bold=True, size=12)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.fill = header_fill
        c.border = thin
    c2.alignment = Alignment(horizontal="right", vertical="center")

    # Überschriften
    r = 3
    headers = ["Name", "Personalnummer", "Gesamtverdienst (€)"]
    for i, h in enumerate(headers, start=start_col):
        cell = sheet.cell(row=r, column=i, value=h)
        cell.font = Font(bold=True, size=12)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = head2_fill
        cell.border = thin

    # Sortierung (absteigend)
    summary_data.sort(key=lambda x: x[2], reverse=True)

    # Daten + Zebra
    zebra = False
    for idx, (name, personalnummer, total) in enumerate(summary_data, start=4):
        zebra = not zebra
        row_fill = row_a_fill if zebra else row_b_fill

        nc = sheet.cell(row=idx, column=start_col, value=name)
        nc.font = Font(size=12)
        nc.alignment = Alignment(horizontal="left", vertical="center")
        nc.fill = row_fill; nc.border = thin

        pn_cell = sheet.cell(row=idx, column=start_col + 1)
        if personalnummer != "Unbekannt":
            pn_cell.value = int(personalnummer)
            pn_cell.number_format = "00000000"
        else:
            pn_cell.value = personalnummer
        pn_cell.font = Font(size=12)
        pn_cell.alignment = Alignment(horizontal="center", vertical="center")
        pn_cell.fill = row_fill; pn_cell.border = thin

        tc = sheet.cell(row=idx, column=start_col + 2, value=float(total))
        tc.number_format = '#,##0.00 €'
        tc.font = Font(size=12)
        tc.alignment = Alignment(horizontal="right", vertical="center")
        tc.fill = row_fill; tc.border = thin

    # Gesamtsumme
    total_row = len(summary_data) + 4
    lab = sheet.cell(row=total_row, column=start_col, value="Gesamtsumme")
    lab.font = Font(bold=True, size=12); lab.fill = total_fill; lab.border = thin
    lab.alignment = Alignment(horizontal="left", vertical="center")

    sumcell = sheet.cell(row=total_row, column=start_col + 2, value=sum(x[2] for x in summary_data))
    sumcell.number_format = '#,##0.00 €'
    sumcell.font = Font(bold=True, size=12); sumcell.fill = total_fill; sumcell.border = thin
    sumcell.alignment = Alignment(horizontal="right", vertical="center")

    # Rahmen
    for row in range(3, total_row + 1):
        for col in range(start_col, start_col + 3):
            cell = sheet.cell(row=row, column=col)
            if cell.value is None:
                cell.value = ""
            cell.border = thin

    # Spaltenbreiten Summary
    widths = {start_col: 28, start_col + 1: 16, start_col + 2: 18}
    for col_idx, width in widths.items():
        sheet.column_dimensions[get_column_letter(col_idx)].width = width

# -------------------------------
# App
# -------------------------------
def format_date_with_german_weekday_wrapper(date):
    # Wrapper nur für Aufrufklarheit
    return format_date_with_german_weekday(pd.to_datetime(date))

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

                # Verdienst (Berechnungslogik bleibt)
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
