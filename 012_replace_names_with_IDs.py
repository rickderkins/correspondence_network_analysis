import pandas as pd
import os


def replace_ids_as_strings():
    # --- 1. Pfad-Definitionen ---
    INPUT_FILE = r'D:\OneDrive - Universität Heidelberg\Studium\Veranstaltungen\19-WiSe25-HIS-MA Masterarbeit\Quellen\Daten\Alt\251127_NODEGOAT_Briefe.csv'

    PERSON_MAPPING_FILE = r'D:\OneDrive - Universität Heidelberg\Studium\Veranstaltungen\19-WiSe25-HIS-MA Masterarbeit\Quellen\Daten\251128_Person-IDs.csv'
    PLACE_MAPPING_FILE = r'D:\OneDrive - Universität Heidelberg\Studium\Veranstaltungen\19-WiSe25-HIS-MA Masterarbeit\Quellen\Daten\251128_Places_with_Geoname-IDs.csv'

    OUTPUT_FILE = r'D:\OneDrive - Universität Heidelberg\Studium\Veranstaltungen\19-WiSe25-HIS-MA Masterarbeit\Quellen\Daten\251128_NODEGOAT_Briefe.csv'

    print("--- Prozess gestartet (String-Modus) ---")

    try:
        # --- 2. Daten laden mit erzwungenem String-Typ ---
        # dtype=str zwingt pandas, ALLE Daten als Text zu lesen.
        # Das verhindert, dass "00123" zu "123" wird oder IDs als Zahlen behandelt werden.

        print("Lade Hauptdatei...")
        df_briefe = pd.read_csv(INPUT_FILE, sep=',', dtype=str)

        print("Lade Mapping-Dateien...")
        df_persons = pd.read_csv(PERSON_MAPPING_FILE, sep=',', dtype=str)
        df_places = pd.read_csv(PLACE_MAPPING_FILE, sep=',', dtype=str)

        # --- 3. Mappings erstellen ---
        # Da wir dtype=str verwendet haben, sind die IDs hier garantiert Strings.
        # Format: dict(zip(Name, ID))

        # Mapping: KORR-NAME -> PERSON-ID
        person_map = dict(zip(df_persons['KORR-NAME'], df_persons['PERSON-ID']))

        # Mapping: ORT-NAME -> ORT-GEONAMES
        place_map = dict(zip(df_places['ORT-NAME'], df_places['ORT-GEONAMES']))

        print(f"Mappings erstellt: {len(person_map)} Personen, {len(place_map)} Orte.")

        # --- 4. Ersetzen der Werte ---
        print("Ersetze ABS-ID und EMP-ID mit Personen-IDs...")
        # Wir mappen die Werte. Unbekannte Werte bleiben erhalten (fillna).
        df_briefe['ABS-ID'] = df_briefe['ABS-ID'].map(person_map).fillna(df_briefe['ABS-ID'])
        df_briefe['EMP-ID'] = df_briefe['EMP-ID'].map(person_map).fillna(df_briefe['EMP-ID'])

        print("Ersetze ABS-GEONAMES und EMP-GEONAMES mit GeoNames-IDs...")
        df_briefe['ABS-GEONAMES'] = df_briefe['ABS-GEONAMES'].map(place_map).fillna(df_briefe['ABS-GEONAMES'])
        df_briefe['EMP-GEONAMES'] = df_briefe['EMP-GEONAMES'].map(place_map).fillna(df_briefe['EMP-GEONAMES'])

        # --- 5. Speichern ---
        print(f"Speichere Datei unter: {OUTPUT_FILE}")

        # Wir speichern ohne Index. Da alles String ist, bleiben Formatierungen erhalten.
        df_briefe.to_csv(OUTPUT_FILE, index=False, sep=',')

        print("✅ Fertig! IDs wurden als Strings verarbeitet und ersetzt.")

    except FileNotFoundError as e:
        print(f"❌ Fehler: Datei nicht gefunden. {e}")
    except KeyError as e:
        print(f"❌ Fehler: Spaltenname nicht gefunden. {e}")
    except Exception as e:
        print(f"❌ Unerwarteter Fehler: {e}")


if __name__ == "__main__":
    replace_ids_as_strings()