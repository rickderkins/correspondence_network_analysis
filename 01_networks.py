import pandas as pd
import folium
import math
import os
import json  # Für die Verarbeitung der GeoJSON-Datei

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
        start_year = int(input("\nFilter 1/2: Geben Sie das Startjahr (z.B. 1890) ein: "))
        end_year = int(input("Filter 1/2: Geben Sie das Endjahr (z.B. 1910) ein: "))
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

# --- ABFRAGE UND ANWENDUNG DES SCHLAGWORT-FILTERS (Filter 2/2) ---

print(
    "\nFilter 2/2: Optional können Sie die Verbindungen zusätzlich nach einem Schlagwort in TAG-FACH, TAG-FUNK oder TAG-INH eingrenzen (z.B. 'archäologie' oder 'literaturversand').")
keyword_filter = input("Schlagwort (leer lassen für alle): ").lower().strip()

# Wenn ein Schlagwort eingegeben wurde, filtere den bereits nach Datum gefilterten DataFrame
if keyword_filter:
    print(f"Wende Schlagwortfilter an: '{keyword_filter}'...")

    # KORREKTUR: Filterung auf alle drei Tag-Spalten (FACH, FUNK, INH) erweitert.
    df_filtered = df_filtered[
        df_filtered['TAG-FACH'].astype(str).str.lower().str.contains(keyword_filter, na=False) |
        df_filtered['TAG-FUNK'].astype(str).str.lower().str.contains(keyword_filter, na=False) |
        df_filtered['TAG-INH'].astype(str).str.lower().str.contains(keyword_filter, na=False)
        ].copy()

    print(f"Nach Schlagwort-Filterung: {len(df_filtered)} Briefe verbleiben.")


# --- SCHRITT 2: KARTE ERSTELLEN UND LINIEN ZEICHNEN (FOLIUM) ---

def get_line_color(row):
    """
    Definiert die Logik für die Einfärbung der Linien basierend auf den Tags.
    """
    tag_fach = str(row['TAG-FACH']).lower()
    tag_funk = str(row['TAG-FUNK']).lower()

    if 'archäologie' in tag_fach or 'archäologe' in tag_funk:
        return 'red'
    if 'geologie' in tag_fach or 'geologe' in tag_funk or 'paläontologe' in tag_funk:
        return 'green'
    if 'medizin' in tag_fach or 'anthropologe' in tag_funk:
        return 'purple'

    return 'blue'


# Den Mittelpunkt der Karte festlegen (z.B. Heidelberg)
map_center = [49.40768, 8.69079]

# Hintergrund: CartoDB DarkMatter für minimalen visuellen Konflikt
m = folium.Map(location=map_center, zoom_start=6, tiles="CartoDB dark_matter")

# --- Laden und Hinzufügen des historischen GeoJSON-Overlays ---

geojson_path = os.path.join(BASE_FOLDER, 'world_1914.geojson')

try:
    with open(geojson_path, 'r', encoding='utf-8') as f:
        historic_borders = json.load(f)

    # GeoJSON als Overlay hinzufügen
    folium.GeoJson(
        historic_borders,
        name='Historische Grenzen (1914)',
        style_function=lambda x: {
            'fillColor': 'none',
            'color': 'white',  # Weiße Linien für guten Kontrast
            'weight': 1.5,
            'dashArray': '5, 5',
            'opacity': 0.9
        },
        tooltip=folium.GeoJsonTooltip(fields=['NAME'], aliases=['Land:'])
    ).add_to(m)

    folium.LayerControl().add_to(m)


except FileNotFoundError:
    print(
        f"\nWARNUNG: GeoJSON-Datei '{os.path.basename(geojson_path)}' nicht gefunden. Bitte stellen Sie sicher, dass sie im Ordner '{BASE_FOLDER}' liegt.")
except Exception as e:
    print(f"\nFEHLER beim Laden des GeoJSON-Overlays: {e}")

print("Karte initialisiert. Zeichne Briefverbindungen...")

# Iterieren durch den GEFILTERTEN DataFrame
for index, row in df_filtered.iterrows():

    start_coords = row['ABS_KOOR_TUPLE']
    end_coords = row['EMP_KOOR_TUPLE']

    if not start_coords or not end_coords:
        continue

    line_color = get_line_color(row)

    popup_text = f"""
    <b>Brief-ID:</b> {row['BRIEF-ID']}<br><br>
    <b>Von:</b> {row['ABS-NAME']} ({row['ABS-ORT']})<br>
    <b>An:</b> {row['EMP-NAME']} ({row['EMP-ORT']})<br>
    <b>Datum:</b> {row['DATUM']}<br>
    <hr>
    <b>Funktion:</b> {row['TAG-FUNK']}<br>
    <b>Fach:</b> {row['TAG-FACH']}
    """

    folium.PolyLine(
        locations=[start_coords, end_coords],
        weight=2,
        color=line_color,
        opacity=0.7,
        popup=folium.Popup(popup_text, max_width=300)
    ).add_to(m)

# --- Speicherung im BASE_FOLDER ---

if keyword_filter:
    output_filename = f"briefnetzwerk_karte_{start_year}-{end_year}_{keyword_filter}.html"
else:
    output_filename = f"briefnetzwerk_karte_{start_year}-{end_year}_alle.html"

output_path = os.path.join(BASE_FOLDER, output_filename)

# Speichern der fertigen Karte als HTML-Datei
m.save(output_path)

print(f"--- FERTIG ---")
print(
    f"Die interaktive Karte (Filter: {start_year}-{end_year}, Schlagwort: {'alle' if not keyword_filter else keyword_filter}) wurde erfolgreich gespeichert unter:\n{output_path}")
print(f"Öffnen Sie diese Datei in einem Webbrowser, um das Ergebnis zu sehen.")