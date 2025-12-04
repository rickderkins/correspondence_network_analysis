import pandas as pd
import os


def extract_and_normalize_tags():
    """
    Fragt nach einer CSV-Datei und einer Spalte, extrahiert eindeutige Tags
    (durch ", " getrennt) und speichert sie in einer neuen CSV-Datei.
    """
    # 📌 Definiere den Basis-Ordnerpfad
    BASE_DIR = r"D:\OneDrive - Universität Heidelberg\Studium\Veranstaltungen\19-WiSe25-HIS-MA Masterarbeit\Quellen\Daten"

    # --- 1. Eingaben vom Benutzer einholen ---
    print("--- Tag-Extraktion aus CSV starten ---")

    # 📝 Fragt nach dem Namen der Eingabedatei
    input_filename = input(f"Bitte geben Sie den Namen der Eingabe-CSV-Datei ein: ")
    input_filepath = os.path.join(BASE_DIR, input_filename)

    # 📝 Fragt nach dem Spaltennamen
    column_name = input("Bitte geben Sie den genauen Namen der Spalte ein, aus der die Tags extrahiert werden sollen: ")

    # --- 2. Datei einlesen und Fehlerbehandlung ---
    if not os.path.exists(input_filepath):
        print(f"❌ Fehler: Datei nicht gefunden unter {input_filepath}")
        return

    try:
        df = pd.read_csv(input_filepath)
        print(f"✅ Datei '{input_filename}' erfolgreich eingelesen.")
    except Exception as e:
        print(f"❌ Fehler beim Lesen der CSV-Datei: {e}")
        return

    if column_name not in df.columns:
        print(
            f"❌ Fehler: Die Spalte '{column_name}' existiert nicht in der Datei. Verfügbare Spalten: {list(df.columns)}")
        return

    # --- 3. Tags extrahieren, aufteilen und normalisieren ---
    print(f"⏳ Extrahiere und bereinige Tags aus Spalte '{column_name}'...")

    # Füllt leere (NaN) Werte in der Spalte mit einem leeren String auf,
    # damit .str.split() keine Fehler verursacht.
    tags_series = df[column_name].fillna('')

    # Teilt jeden String in eine Liste von Tags auf (Trenner ist ", ")
    # .explode() macht aus jeder Liste eine einzelne Zeile pro Tag
    # .str.strip() entfernt führende/nachfolgende Leerzeichen von jedem Tag
    all_tags = tags_series.str.split(', ').explode().str.strip()

    # Entfernt leere Strings, die durch das .fillna('') oder leere Einträge im Original entstehen
    unique_tags = all_tags[all_tags != ''].drop_duplicates()

    print(f"✨ {len(unique_tags)} eindeutige Tags gefunden.")

    # --- 4. Ergebnis-DataFrame erstellen und speichern ---

    # Erstellt einen neuen DataFrame mit der gewünschten Spalte und dem Header
    output_df = pd.DataFrame({column_name: unique_tags.sort_values().tolist()})

    # 📝 Definiere den Ausgabedateinamen
    # Ersetze ungültige Zeichen im Spaltennamen für den Dateinamen
    safe_column_name = "".join([c for c in column_name if c.isalnum() or c in ('_', '-')]).rstrip()
    output_filename = f"251204_NODEGOAT_{safe_column_name}.csv"
    output_filepath = os.path.join(BASE_DIR, output_filename)

    # Speichert das Ergebnis in einer neuen CSV-Datei
    # index=False verhindert, dass der Pandas-Index als Spalte gespeichert wird
    output_df.to_csv(output_filepath, index=False)

    print(f"🎉 Erfolg! Die eindeutigen Tags wurden gespeichert unter:")
    print(f"    {output_filepath}")
    print("--- Tag-Extraktion abgeschlossen ---")


# Starte die Funktion
if __name__ == "__main__":
    # Stelle sicher, dass pandas installiert ist: pip install pandas
    extract_and_normalize_tags()