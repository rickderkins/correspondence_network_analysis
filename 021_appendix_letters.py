import csv
import os
from docx import Document
from docx.oxml import OxmlElement
from copy import deepcopy


def erstelle_brief_anhang():
    # Absolute Pfade
    input_csv = r"D:\heiBOX\Seafile\Masterarbeit_Ablage\daten_heibox\260115_NODEGOAT_Briefe.csv"
    template_path = r"D:\heiBOX\Seafile\Masterarbeit_Ablage\daten_heibox\template_appendix_letter.docx"
    output_docx = r"D:\heiBOX\Seafile\Masterarbeit_Ablage\daten_heibox\Anhang_Test.docx"

    if not os.path.exists(template_path) or not os.path.exists(input_csv):
        print("Fehler: Pfade prüfen!")
        return

    doc = Document(template_path)

    # Vorlage extrahieren
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

            # Vorlage im Dokument löschen
            template_para._element.getparent().remove(template_para._element)
            template_table._element.getparent().remove(template_table._element)

            # Zähler initialisieren
            for i, row in enumerate(reader, start=1):
                # 1. TITEL-ABSATZ einfügen
                new_p_element = deepcopy(template_para_xml)
                doc._element.body.append(new_p_element)

                aktiver_absatz = doc.paragraphs[-1]
                if "BRIEF-TITEL" in aktiver_absatz.text:
                    wert = str(row.get("BRIEF-TITEL", "")).strip()
                    # Nummerierung hinzufügen: "1. Titel-Inhalt"
                    neuer_text = f"{i}. {wert}"
                    aktiver_absatz.text = aktiver_absatz.text.replace("BRIEF-TITEL", neuer_text)

                # 2. TABELLE einfügen
                new_tbl_element = deepcopy(template_table_xml)
                doc._element.body.append(new_tbl_element)

                aktuelle_tabelle = doc.tables[-1]
                for zeile in aktuelle_tabelle.rows:
                    for zelle in zeile.cells:
                        for spalte in spalten_liste:
                            if spalte in zelle.text:
                                wert = str(row.get(spalte, "")).strip()
                                zelle.text = zelle.text.replace(spalte, wert)

                # 3. LEERZEILE einfügen
                doc._element.body.append(OxmlElement('w:p'))

        doc.save(output_docx)
        print(f"Erfolg! {i} Briefe wurden nummeriert und in '{output_docx}' gespeichert.")

    except Exception as e:
        print(f"Ein Fehler ist aufgetreten: {e}")


if __name__ == "__main__":
    erstelle_brief_anhang()