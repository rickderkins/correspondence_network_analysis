# dieses Skript dient der Übertragung der Tags zu Funktion und Fach von der Personen- in die Korrepondenztabelle

import pandas as pd
import numpy as np
import sys
import os
import logging
from datetime import datetime  # <--- Hinzugefügt für das Datum

# --- 1. Konfiguration ---
# Der Ordner, in dem Ihre Dateien liegen
folder_path = r'D:\OneDrive - Universität Heidelberg\Studium\Veranstaltungen\19-WiSe25-HIS-MA Masterarbeit\Quellen\Daten'

# Name für die Log-Datei
log_file_name = 'verarbeitung.log'
log_file_path = os.path.join(folder_path, log_file_name)

# --- Logging Konfiguration ---
# (Unverändert)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    handlers=[
        logging.FileHandler(log_file_path, encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)

# --- Dateinamen abfragen ---
print("\n--- Skript zur Datenanreicherung ---")
print("Bitte geben Sie die Namen der Input-Dateien an.")
print(f"Die Dateien werden im Ordner '{folder_path}' gesucht.")

file_korrespondenz_name = input("1. Name der Korrespondenz-Datei (z.B. 251108_Gesamtkorrespondenz.csv): ")
file_personen_name = input("2. Name der Personenliste-Datei (z.B. 251108_Personenliste.csv): ")

# --- NEU: Ausgabedatei automatisch generieren ---
# Hole das heutige Datum
heute = datetime.now()
# Formatiere das Datum als YYMMDD (z.B. 251109)
datum_string = heute.strftime("%y%m%d")
# Erstelle den Dateinamen
file_output_name = f"{datum_string}_Gesamtkorrespondenz.csv"

print(f"\nDer Name der Ausgabedatei wird automatisch gesetzt auf: '{file_output_name}'")
print("----------------------------------------------------------")

# Komplette Pfade basierend auf den Eingaben erstellen
file_korrespondenz = os.path.join(folder_path, file_korrespondenz_name)
file_personen = os.path.join(folder_path, file_personen_name)
file_output = os.path.join(folder_path, file_output_name)

# Der Name, der ignoriert werden soll
key_person = "Schoetensack, Otto"

# --- Start der Verarbeitung (Logging) ---
logging.info("--- Skript gestartet ---")
logging.info(f"Lese-Ordner: {folder_path}")
logging.info(f"Lade Korrespondenzdatei: '{file_korrespondenz_name}'")
logging.info(f"Lade Personendatei: '{file_personen_name}'")
logging.info(f"Zieldatei wird automatisch generiert: '{file_output_name}'")

try:
    # --- 2. Daten laden ---
    logging.info(f"Lade '{file_korrespondenz_name}'...")
    df_korr = pd.read_csv(file_korrespondenz, dtype=str)

    logging.info(f"Lade '{file_personen_name}'...")
    df_person = pd.read_csv(file_personen, dtype=str)

    logging.info("Dateien erfolgreich geladen.")

    # --- 3. Nachschlage-Maps (Lookup) erstellen ---
    logging.info("Erstelle Nachschlage-Verzeichnisse aus der Personenliste...")
    funk_map = df_person.set_index('KORR-NAME')['KORR-FUNK'].fillna('').squeeze()
    fach_map = df_person.set_index('KORR-NAME')['KORR-FACH'].fillna('').squeeze()
    logging.info("Nachschlage-Verzeichnisse erstellt.")

    # --- 4. Korrespondenzpartner identifizieren ---
    logging.info("Identifiziere Korrespondenzpartner für jeden Brief...")
    df_korr['partner_name'] = np.where(
        df_korr['ABS-NAME'] == key_person,  # Bedingung
        df_korr['EMP-NAME'],  # Wert, wenn Bedingung wahr
        df_korr['ABS-NAME']  # Wert, wenn Bedingung falsch
    )
    logging.info("Partner identifiziert.")

    # --- 5. Daten anreichern (Mapping) ---
    logging.info("Reichere Daten an: Fülle TAG-FUNK und TAG-FACH...")
    df_korr['TAG-FUNK'] = df_korr['partner_name'].map(funk_map)
    df_korr['TAG-FACH'] = df_korr['partner_name'].map(fach_map)

    df_korr['TAG-FUNK'] = df_korr['TAG-FUNK'].fillna('')
    df_korr['TAG-FACH'] = df_korr['TAG-FACH'].fillna('')
    logging.info("Datenanreicherung abgeschlossen.")

    # --- 6. Aufräumen ---
    df_korr = df_korr.drop(columns=['partner_name'])
    logging.info("Hilfsspalte 'partner_name' entfernt.")

    # --- 7. Neue CSV-Datei speichern ---
    logging.info(f"Speichere angereicherte Daten in '{file_output_name}'...")
    df_korr.to_csv(file_output, index=False, encoding='utf-8-sig')

    logging.info("--- ERFOLGREICH ---")
    logging.info(f"Verarbeitung abgeschlossen. Die neue Datei wurde als '{file_output_name}' gespeichert.")
    logging.info(f"Ein detailliertes Protokoll finden Sie in: '{log_file_path}'")


# --- Fehlerbehandlung mit Logging ---
except FileNotFoundError as e:
    logging.error("--- FEHLER: Datei nicht gefunden ---")
    logging.error(f"Konnte die Datei '{e.filename}' nicht finden.")
    logging.error("Bitte überprüfen Sie die Schreibweise und stellen Sie sicher, dass die Datei im Ordner existiert.")

except KeyError as e:
    logging.error("--- FEHLER: Spalte nicht gefunden ---")
    logging.error(f"Die Spalte {e} wurde in einer der CSV-Dateien nicht gefunden.")
    logging.error("Bitte überprüfen Sie die Spaltennamen in Ihren Dateien (Groß-/Kleinschreibung beachten!).")

except Exception as e:
    logging.exception("--- Ein unerwarteter Fehler ist aufgetreten ---")

logging.info("--- Skript beendet ---")