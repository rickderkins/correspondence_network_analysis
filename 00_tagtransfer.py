import pandas as pd
import os


def daten_anreichern():
    # 1. Pfad und Dateinamen definieren
    # Das 'r' vor den Anführungszeichen ist wichtig, damit die Backslashes (\) korrekt gelesen werden.
    base_path = r"D:\OneDrive - Universität Heidelberg\Studium\Veranstaltungen\19-WiSe25-HIS-MA Masterarbeit\Quellen\Daten"

    # Wir nutzen os.path.join, um den Ordnerpfad sauber mit den Dateinamen zu verbinden
    file_personen = os.path.join(base_path, '251204_NODEGOAT_Personen.csv')
    file_briefe = os.path.join(base_path, '251204_NODEGOAT_Briefe.csv')
    file_output = os.path.join(base_path, '251204_NODEGOAT_Briefe_angereichert.csv')

    print(f"Arbeitsverzeichnis: {base_path}")
    print("Lade Daten...")

    # Prüfen, ob Dateien existieren
    if not os.path.exists(file_personen) or not os.path.exists(file_briefe):
        print(f"Fehler: Die Dateien wurden unter dem angegebenen Pfad nicht gefunden.")
        return

    # Daten einlesen
    # HINWEIS: Falls deine CSVs Semikolons (;) statt Kommas als Trenner nutzen,
    # ändere unten sep=',' zu sep=';'
    try:
        df_pers = pd.read_csv(file_personen, sep=',', dtype=str, encoding='utf-8')
        df_briefe = pd.read_csv(file_briefe, sep=',', dtype=str, encoding='utf-8')
    except UnicodeDecodeError:
        # Fallback, falls die Datei nicht utf-8 codiert ist (z.B. bei Excel-Exporten oft 'latin1')
        print("UTF-8 Fehler, versuche 'latin1' Kodierung...")
        df_pers = pd.read_csv(file_personen, sep=',', dtype=str, encoding='latin1')
        df_briefe = pd.read_csv(file_briefe, sep=',', dtype=str, encoding='latin1')

    # NaN (leere Felder) durch leere Strings ersetzen
    df_pers = df_pers.fillna('')
    df_briefe = df_briefe.fillna('')

    print("Erstelle Such-Index für Personen...")

    # 2. Personen-Daten in ein Dictionary umwandeln
    personen_lookup = {}
    for index, row in df_pers.iterrows():
        p_id = row['PERSON-ID']
        personen_lookup[p_id] = {
            'TAG-TAET': row['TAG-TAET'],
            'TAG-FACH': row['TAG-FACH'],
            'TAG-INST': row['TAG-INST']
        }

    print("Bearbeite Briefe (das kann einen Moment dauern)...")

    # 3. Funktion für das Update der Zeilen
    def update_row(row):
        target_id = ''

        # Logik: Wenn ABS-ID nicht 069-SO ist, nehmen wir diese. Sonst EMP-ID.
        if row['ABS-ID'] != '069-SO':
            target_id = row['ABS-ID']
        else:
            target_id = row['EMP-ID']

        # Nachschlagen
        if target_id in personen_lookup:
            match = personen_lookup[target_id]
            row['NONOS-TAET'] = match['TAG-TAET']
            row['NONOS-FACH'] = match['TAG-FACH']
            row['NONOS-INST'] = match['TAG-INST']

        return row

    # Die Funktion anwenden
    df_briefe_neu = df_briefe.apply(update_row, axis=1)

    # 4. Speichern
    print(f"Speichere neue Datei: {file_output}")
    df_briefe_neu.to_csv(file_output, index=False, sep=',', encoding='utf-8')
    print("Fertig!")


if __name__ == "__main__":
    daten_anreichern()