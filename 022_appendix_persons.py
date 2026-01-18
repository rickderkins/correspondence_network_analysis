import csv
import os
from docx import Document
from docx.oxml import OxmlElement
from copy import deepcopy
from datetime import datetime


def erstelle_personen_anhang():
    # 1. Zeitstempel generieren (Format: YYMMDD_HHMM)
    zeitstempel = datetime.now().strftime("%y%m%d_%H%M")
    dateiname = f"{zeitstempel}_Anhang_Personen.docx"

    # 2. Pfade definieren
    input_csv = r"D:\heiBOX\Seafile\Masterarbeit_Ablage\daten_heibox\260117_PERSONEN.csv"
    template_path = r"D:\heiBOX\Seafile\Masterarbeit_Ablage\anhang_heibox\template_appendix_person.docx"
    output_dir = r"D:\heiBOX\Seafile\Masterarbeit_Ablage\anhang_heibox"
    output_docx = os.path.join(output_dir, dateiname)

    # Sicherstellen, dass der Output-Ordner existiert
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    if not os.path.exists(template_path) or not os.path.exists(input_csv):
        print("Fehler: Pfade prüfen!")
        return

    # 3. Dokument laden und Vorlage extrahieren
    doc = Document(template_path)
    template_para = doc.paragraphs[0]
    template_table = doc.tables[0]

    template_para_xml = deepcopy(template_para._element)
    template_table_xml = deepcopy(template_table._element)

    spalten_liste = [
        "KORR-NACHNAME", "KORR-VORNAME", "KORR-BERU",
        "TAG-TAET", "TAG-FACH", "TAG-INST",
        "REF-1", "REF-2", "REF-3", "REF-4"
    ]

    try:
        with open(input_csv, mode='r', encoding='utf-8-sig') as csvfile:
            sample = csvfile.read(4096)
            csvfile.seek(0)
            dialect = csv.Sniffer().sniff(sample)
            reader = csv.DictReader(csvfile, dialect=dialect)

            # Vorlage im Dokument löschen
            template_para._element.getparent().remove(template_para._element)
            template_table._element.getparent().remove(template_table._element)

            i = 0
            for i, row in enumerate(reader, start=1):
                # A) Absatz einfügen & Platzhalter ersetzen
                new_p_element = deepcopy(template_para_xml)
                doc._element.body.append(new_p_element)
                aktiver_absatz = doc.paragraphs[-1]

                for spalte in spalten_liste:
                    if spalte in aktiver_absatz.text:
                        wert = str(row.get(spalte, "")).strip()
                        aktiver_absatz.text = aktiver_absatz.text.replace(spalte, wert)

                # B) Tabelle einfügen & Platzhalter ersetzen
                new_tbl_element = deepcopy(template_table_xml)
                doc._element.body.append(new_tbl_element)
                aktuelle_tabelle = doc.tables[-1]

                for zeile in aktuelle_tabelle.rows:
                    for zelle in zeile.cells:
                        for spalte in spalten_liste:
                            if spalte in zelle.text:
                                wert = str(row.get(spalte, "")).strip()
                                zelle.text = zelle.text.replace(spalte, wert)

                # C) Leerzeile einfügen
                doc._element.body.append(OxmlElement('w:p'))

        # 4. Speichern mit dem neuen Dateinamen
        doc.save(output_docx)
        print(f"Erfolg! {i} Personen wurden verarbeitet.")
        print(f"Datei gespeichert als: {dateiname}")

    except Exception as e:
        print(f"Ein Fehler ist aufgetreten: {e}")


if __name__ == "__main__":
    erstelle_personen_anhang()