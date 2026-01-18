import csv
import os
from datetime import datetime
from docx import Document
from docx.oxml import OxmlElement
from copy import deepcopy


def erstelle_brief_anhang():
    # 1. Pfade definieren
    input_csv = r"D:\heiBOX\Seafile\Masterarbeit_Ablage\daten_heibox\260117_BRIEFE.csv"
    template_path = r"D:\heiBOX\Seafile\Masterarbeit_Ablage\anhang_heibox\template_appendix_letter.docx"
    output_folder = r"D:\heiBOX\Seafile\Masterarbeit_Ablage\anhang_heibox"

    # Zeitstempel für den Dateinamen generieren (YYMMDD_HHMM)
    zeitstempel = datetime.now().strftime("%y%m%d_%H%M")
    output_filename = f"{zeitstempel}_Anhang_Briefe.docx"
    output_docx = os.path.join(output_folder, output_filename)

    # Validierung der Pfade
    if not os.path.exists(template_path):
        print(f"Fehler: Vorlage nicht gefunden: {template_path}")
        return
    if not os.path.exists(input_csv):
        print(f"Fehler: CSV nicht gefunden: {input_csv}")
        return
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Dokument laden
    doc = Document(template_path)

    # Vorlage extrahieren (Titel-Absatz und Tabelle)
    template_para = doc.paragraphs[0]
    template_table = doc.tables[0]

    template_para_xml = deepcopy(template_para._element)
    template_table_xml = deepcopy(template_table._element)

    spalten_liste = [
        "BRIEF-TITEL", "ABS-VORNAME", "ABS-NACHNAME", "ABS-ORT",
        "EMP-VORNAME", "EMP-NACHNAME", "EMP-ORT",
        "DATUM", "ARCHIV", "TRANSK", "TAG-FUNK", "TAG-THEMA"
    ]

    try:
        with open(input_csv, mode='r', encoding='utf-8-sig') as csvfile:
            sample = csvfile.read(4096)
            csvfile.seek(0)
            dialect = csv.Sniffer().sniff(sample)
            reader = csv.DictReader(csvfile, dialect=dialect)

            # Vorlage im Dokument löschen, um sauber neu aufzubauen
            template_para._element.getparent().remove(template_para._element)
            template_table._element.getparent().remove(template_table._element)

            brief_zaehler = 0
            for i, row in enumerate(reader, start=1):
                # --- A: TITEL-ABSATZ ---
                new_p_element = deepcopy(template_para_xml)
                doc._element.body.append(new_p_element)

                aktiver_absatz = doc.paragraphs[-1]
                if "BRIEF-TITEL" in aktiver_absatz.text:
                    wert = str(row.get("BRIEF-TITEL", "")).strip()
                    neuer_text = f"{i}. {wert}"
                    aktiver_absatz.text = aktiver_absatz.text.replace("BRIEF-TITEL", neuer_text)

                # --- B: TABELLE ---
                new_tbl_element = deepcopy(template_table_xml)
                doc._element.body.append(new_tbl_element)

                aktuelle_tabelle = doc.tables[-1]
                for zeile in aktuelle_tabelle.rows:
                    for zelle in zeile.cells:
                        for spalte in spalten_liste:
                            if spalte in zelle.text:
                                wert = str(row.get(spalte, "")).strip()
                                zelle.text = zelle.text.replace(spalte, wert)

                # --- C: LEERZEILE ---
                doc._element.body.append(OxmlElement('w:p'))
                brief_zaehler = i

        # Speichern des Dokuments
        doc.save(output_docx)
        print(f"Erfolg! {brief_zaehler} Briefe verarbeitet.")
        print(f"Datei gespeichert unter: {output_docx}")

    except Exception as e:
        print(f"Ein Fehler ist aufgetreten: {e}")


if __name__ == "__main__":
    erstelle_brief_anhang()