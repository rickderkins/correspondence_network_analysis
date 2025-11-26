import pandas as pd
import folium
import math
import os
import json

print("Starte Skript...")


# --- SCHRITT 1: DATEN LADEN UND VORBEREITEN (PANDAS) ---

def parse_coords(coord_str):
    """
    Konvertiert einen Koordinaten-String "lat, lon" in ein Tupel (lat, lon).
    Behandelt fehlende, ungültige und reine Text-Werte robust.
    """
    if pd.isna(coord_str):
        return None

    cleaned_str = str(coord_str).replace(' ', '')

    if ',' not in cleaned_str:
        return None

    try:
        parts = cleaned_str.split(',')
        if len(parts) != 2:
            return None

        lat = float(parts[0])
        lon = float(parts[1])

        if -90 <= lat <= 90 and -180 <= lon <= 180:
            return (lat, lon)
        else:
            return None

    except ValueError:
        return None
    except Exception:
        return None


# --- Fester Ordnerpfad ---
BASE_FOLDER = r"D:\OneDrive - Universität Heidelberg\Studium\Veranstaltungen\19-WiSe25-HIS-MA Masterarbeit\Quellen\Daten"

print(f"Das Skript sucht im folgenden Ordner nach Dateien:\n{BASE_FOLDER}\n")
print("Bitte geben Sie den Dateinamen Ihrer CSV-Datei ein (inkl. .csv Endung).")
file_name = input("Dateiname: ")

# --- Kombinieren von Ordner und Dateiname ---
file_path = os.path.join(BASE_FOLDER, file_name)

# Laden der CSV-Datei
try:
    df = pd.read_csv(file_path)
    print(f"CSV-Datei '{file_path}' erfolgreich geladen. {len(df)} Zeilen gefunden.")

    df['DATUM_DATE'] = pd.to_datetime(df['DATUM'], errors='coerce')

except FileNotFoundError:
    print(f"--- FEHLER ---")
    print(f"Die Datei '{file_name}' wurde im Ordner '{BASE_FOLDER}' nicht gefunden.")
    print("Bitte überprüfen Sie die Schreibweise des Dateinamens und des Pfades.")
    exit()
except Exception as e:
    print(f"Ein unerwarteter Fehler ist beim Laden der Datei aufgetreten: {e}")
    exit()

# Anwenden der robusteren Parsing-Funktion auf die Koordinatenspalten
df['ABS_KOOR_TUPLE'] = df['ABS-KOOR'].apply(parse_coords)
df['EMP_KOOR_TUPLE'] = df['EMP-KOOR'].apply(parse_coords)

# --- ABFRAGE UND ANWENDUNG DES DATUMS-FILTERS ---

while True:
    try:
        start_year = int(input("\nFilter: Geben Sie das Startjahr (z.B. 1890) ein: "))
        end_year = int(input("Filter: Geben Sie das Endjahr (z.B. 1910) ein: "))
        if start_year <= end_year and start_year > 1800:
            break
        else:
            print("Ungültiger Zeitraum. Das Startjahr muss vor oder gleich dem Endjahr liegen.")
    except ValueError:
        print("Ungültige Eingabe. Bitte geben Sie eine ganze Zahl (Jahr) ein.")

print(f"Wende Datumsfilter an: {start_year} bis {end_year}...")
df_filtered = df[(df['DATUM_DATE'].dt.year >= start_year) &
                 (df['DATUM_DATE'].dt.year <= end_year)].copy()

print(f"Nach Datumsfilterung: {len(df_filtered)} Briefe verbleiben.")

# --- SCHRITT 2: KARTE ERSTELLEN UND GRUPPEN DEFINIEREN (FOLIUM) ---

# Den Mittelpunkt der Karte festlegen (z.B. Heidelberg)
map_center = [49.40768, 8.69079]

# Hintergrund: Standard-Basiskarte
m = folium.Map(location=map_center, zoom_start=6, tiles="CartoDB positron")

# --- DEFINITION DER INTERAKTIVEN LAYER ---

# Layer für Inhaltsfilter
fg_object = folium.FeatureGroup(name='1. Inhalt: Objekt-Diskussion (Rot)').add_to(m)
fg_literature = folium.FeatureGroup(name='2. Inhalt: Literatur-Austausch (Grün)').add_to(m)
fg_other = folium.FeatureGroup(name='3. Inhalt: Sonstige (Blau)').add_to(m)

# NEU: Layer für Personenfilter (Sender vs. Empfänger)
fg_schoetensack_sender = folium.FeatureGroup(name='4. Person: S. als Absender (Orange)').add_to(m)
fg_schoetensack_recipient = folium.FeatureGroup(name='5. Person: S. als Empfänger (Teal)').add_to(m)

print("Karte initialisiert. Zeichne Briefverbindungen und ordne sie Layern zu...")

# Iterieren durch den GEFILTERTEN DataFrame
for index, row in df_filtered.iterrows():

    start_coords = row['ABS_KOOR_TUPLE']
    end_coords = row['EMP_KOOR_TUPLE']

    if not start_coords or not end_coords:
        continue

        # Normalisierung der Spalten für die Prüfung
    tag_inh = str(row['TAG-INH']).lower()

    # NEU: Bereinigung (Kleinschreibung, Leerzeichen entfernen) für exakten Abgleich
    abs_name_clean = str(row['ABS-NAME']).lower().strip()
    emp_name_clean = str(row['EMP-NAME']).lower().strip()

    # Popup-Text vorbereiten
    popup_text = f"""
    <b>Brief-ID:</b> {row['BRIEF-ID']}<br><br>
    <b>Von:</b> {row['ABS-NAME']} ({row['ABS-ORT']})<br>
    <b>An:</b> {row['EMP-NAME']} ({row['EMP-ORT']})<br>
    <b>Datum:</b> {row['DATUM']}<br>
    <hr>
    <b>Funktion:</b> {row['TAG-FUNK']}<br>
    <b>Fach:</b> {row['TAG-FACH']}<br>
    <b>Inhalt:</b> {row['TAG-INH']}<br>
    """

    assigned_to_content_group = False

    # ZUWEISUNG NACH INHALT (TAG-INH)

    # 1. Objekt-Diskussion (ROT)
    if 'objekt' in tag_inh:
        folium.PolyLine(
            locations=[start_coords, end_coords],
            weight=2,
            color='red',
            opacity=0.7,
            popup=folium.Popup(popup_text, max_width=300)
        ).add_to(fg_object)
        assigned_to_content_group = True

    # 2. Literatur-Austausch (GRÜN)
    if 'literatur' in tag_inh:
        folium.PolyLine(
            locations=[start_coords, end_coords],
            weight=2,
            color='green',
            opacity=0.7,
            popup=folium.Popup(popup_text, max_width=300)
        ).add_to(fg_literature)
        assigned_to_content_group = True

    # 3. Sonstige Inhalte (BLAU)
    if not assigned_to_content_group:
        folium.PolyLine(
            locations=[start_coords, end_coords],
            weight=2,
            color='blue',
            opacity=0.7,
            popup=folium.Popup(popup_text, max_width=300)
        ).add_to(fg_other)

    # ZUWEISUNG NACH PERSON (EXAKTE PRÜFUNG)

    # 4. Schoetensack als ABSENDER (ORANGE)
    if abs_name_clean == 'schoetensack, otto':
        folium.PolyLine(
            locations=[start_coords, end_coords],
            weight=3,  # Dickere Linie zur Hervorhebung
            color='orange',
            opacity=0.9,
            popup=folium.Popup(popup_text, max_width=300)
        ).add_to(fg_schoetensack_sender)

    # 5. Schoetensack als EMPFÄNGER (TEAL)
    if emp_name_clean == 'schoetensack, otto':
        folium.PolyLine(
            locations=[start_coords, end_coords],
            weight=3,  # Dickere Linie zur Hervorhebung
            color='teal',
            opacity=0.9,
            popup=folium.Popup(popup_text, max_width=300)
        ).add_to(fg_schoetensack_recipient)

# --- Layer-Steuerung hinzufügen ---

# LayerControl ermöglicht dem Benutzer, die FeatureGroups im Browser ein- und auszuschalten.
folium.LayerControl().add_to(m)

# --- Speicherung im BASE_FOLDER ---

output_filename = f"briefnetzwerk_karte_abs_empf_{start_year}-{end_year}.html"
output_path = os.path.join(BASE_FOLDER, output_filename)

# Speichern der fertigen Karte als HTML-Datei
m.save(output_path)

print(f"--- FERTIG ---")
print(
    f"Die interaktive Karte mit detaillierter Schoetensack-Filterung (Filter: {start_year}-{end_year}) wurde erfolgreich gespeichert unter:\n{output_path}")
print(f"Sie können nun die Layer 'S. als Absender (Orange)' und 'S. als Empfänger (Teal)' unabhängig steuern.")