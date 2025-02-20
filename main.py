#!/usr/bin/env python3
# main.py
"""
Aplicație GUI locală în Python/PyQt5 (în limba română), folosind PyMuPDF (fitz):
1) "Selectează PDF-uri"
2) "Descoperă câmpuri": extrage atașamente IN-MEMORY (fără să le scrie pe disk),
   parsează fișierele XML și adună tag-urile unice. Evităm duplicarea fișierelor.
3) Mapare "Tag XML" -> "Coloană CSV"
4) "Salvează maparea" -> scrie mapping_config.json
5) "Procesează -> CSV":
   - Creează subfolder "extracted_xml/<timestamp>"
   - Salvează atașamentele .xml acolo
   - Generează "output_<timestamp>.csv"
   - Dacă un tag apare de mai multe ori, generăm rânduri separate (și duplicăm datele celorlalte taguri).

Pași de rulare:
  python main.py

Pentru a crea un .exe Windows cu PyInstaller:
  pip install pyinstaller
  pyinstaller --noconfirm --onefile --windowed main.py
"""

import sys
import os
import re
import json
import csv
import datetime
from typing import Set, Dict, List

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QMessageBox, QTableWidget,
    QTableWidgetItem, QHeaderView, QAction
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon

import fitz  # PyMuPDF
from lxml import etree


########################################################################
# 0) Funcții de ajutor
########################################################################

def sanitize_filename(name: str) -> str:
    """Elimină caractere invalide (\\/:*?"<>|\r\n) dintr-un nume de fișier (Windows)."""
    return re.sub(r'[\\/:*?"<>|\r\n]+', '_', name)

def este_xml_in_memory(data: bytes) -> bool:
    """Verifică dacă un buffer (bytes) reprezintă un fișier XML valid (prin lxml)."""
    try:
        etree.fromstring(data)
        return True
    except:
        return False

def parse_xml_fields_in_memory(data: bytes) -> Set[str]:
    """
    Parsează un fișier XML (din memorie) și returnează setul de tag-uri găsite.
    """
    fields = set()
    try:
        root = etree.fromstring(data)
        for elem in root.iter():
            fields.add(elem.tag)
    except Exception as e:
        print(f"EROARE parse_xml_fields_in_memory: {e}")
    return fields

def parse_xml_and_extract_rows(xml_file: str, field_mapping: Dict[str, str]) -> List[Dict[str, str]]:
    """
    Citește un fișier XML de pe disk și returnează O LISTĂ DE RÂNDURI (fiecare rând e un dict).
    
    *Cerință*: Dacă un tag apare de N ori, generăm N rânduri și duplicăm valorile unice.
    - Pas 1: colectăm toate aparițiile pentru fiecare tag (din field_mapping).
    - Pas 2: calculăm 'max_count' = cel mai mare număr de apariții printre tag-urile mapate.
    - Pas 3: generăm rânduri. Pe rândul i, pentru un tag care are M apariții, 
      folosim apariția min(i, M-1) (adică repetăm ultima apariție dacă i >= M).
    """
    rows = []
    if not os.path.isfile(xml_file):
        return rows

    # 0) parse XML
    try:
        tree = etree.parse(xml_file)
        root = tree.getroot()
    except Exception as e:
        print(f"Eroare la parsearea {xml_file}: {e}")
        return rows

    # 1) strângem toate aparițiile: tag_occurrences[tag_xml] = [val1, val2, ...]
    tag_occurrences = {}
    for xml_tag, csv_header in field_mapping.items():
        elements = root.findall('.//' + xml_tag)
        # colectăm textul (strip). Dacă nu există, e "", etc.
        values = []
        for el in elements:
            if el.text is not None:
                values.append(el.text.strip())
        if not values:
            # nimic găsit => punem un singur element, ex. "" => user vrea să duplăm oricum
            values = [""]
        tag_occurrences[xml_tag] = values

    # 2) calculăm max_count
    max_count = max(len(vals) for vals in tag_occurrences.values()) if tag_occurrences else 0
    if max_count == 0:
        return rows

    # 3) generăm rândurile
    for i in range(max_count):
        row_data = {}
        # inițializăm cu field_mapping values => CSV headers
        for xml_tag, csv_header in field_mapping.items():
            occurrences = tag_occurrences[xml_tag]
            # index = min(i, len(occurrences)-1)
            idx = i if i < len(occurrences) else len(occurrences)-1
            row_data[csv_header] = occurrences[idx]
        rows.append(row_data)

    return rows


########################################################################
# 1) "Descoperă câmpuri" IN-MEMORY
########################################################################

def discover_xml_fields_in_memory(pdf_path: str) -> Set[str]:
    """
    Deschide PDF-ul, extrage atașamente (document-level + adnotări) DOAR ÎN MEMORIE,
    parsează fiecare ca XML (dacă valid), și returnează toate tag-urile găsite (set).
    """
    fields = set()
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        print(f"Eroare la deschiderea {pdf_path}: {e}")
        return fields

    # 1) Document-level attachments
    emb_count = doc.embfile_count()
    for i in range(emb_count):
        info = doc.embfile_info(i)
        attach_name = info.get("filename", f"attachment_{i}.bin")
        data = doc.embfile_get(i)  # bytes

        # test if .xml or parse directly
        if attach_name.lower().endswith('.xml') or este_xml_in_memory(data):
            fields_in_file = parse_xml_fields_in_memory(data)
            fields.update(fields_in_file)

    # 2) Annotation-based attachments
    for page_index in range(len(doc)):
        page = doc.load_page(page_index)
        annots = page.annots()
        if not annots:
            continue

        for annot in annots:
            if annot.type[0] == fitz.PDF_ANNOT_FILEATTACHMENT:
                data = annot.file_get()
                # check if .xml or parse
                finfo = annot.file_info()
                fname = finfo.get('filename', f"page{page_index}.bin")
                if fname.lower().endswith('.xml') or este_xml_in_memory(data):
                    fields_in_file = parse_xml_fields_in_memory(data)
                    fields.update(fields_in_file)

    doc.close()
    return fields


########################################################################
# 2) Extrage "pe disk" doar la PASUL "Procesează -> CSV"
########################################################################

def extract_xml_attachments_to_disk(pdf_path: str, output_dir: str) -> List[str]:
    """
    Similar cu extract_xml_attachments, dar extrage fișiere .xml pe disk.
    Returnează calea fișierelor .xml. Folosit doar la "Procesează -> CSV".
    """
    if not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)

    doc = fitz.open(pdf_path)
    extracted_files = []

    # 1) Document-level
    emb_count = doc.embfile_count()
    for i in range(emb_count):
        info = doc.embfile_info(i)
        attach_name = info.get("filename", f"attachment_{i}.bin")
        data = doc.embfile_get(i)

        safe_pdf = sanitize_filename(os.path.basename(pdf_path))
        safe_name = sanitize_filename(attach_name)
        out_path = os.path.join(output_dir, f"{safe_pdf}_{safe_name}")

        with open(out_path, 'wb') as f:
            f.write(data)

        if safe_name.lower().endswith('.xml'):
            extracted_files.append(out_path)
        else:
            # fallback: dacă e XML valid
            if este_xml_in_memory(data):
                new_xml = out_path + ".xml"
                os.rename(out_path, new_xml)
                extracted_files.append(new_xml)
            else:
                os.remove(out_path)

    # 2) Annotation-based (paperclip)
    for page_index in range(len(doc)):
        page = doc.load_page(page_index)
        annots = page.annots()
        if not annots:
            continue

        for annot in annots:
            if annot.type[0] == fitz.PDF_ANNOT_FILEATTACHMENT:
                data = annot.file_get()
                finfo = annot.file_info()
                attach_name = finfo.get('filename', f"page{page_index}.bin")

                safe_pdf = sanitize_filename(os.path.basename(pdf_path))
                safe_name = sanitize_filename(attach_name)
                out_path = os.path.join(output_dir, f"{safe_pdf}_{safe_name}")

                with open(out_path, 'wb') as f:
                    f.write(data)

                if safe_name.lower().endswith('.xml'):
                    extracted_files.append(out_path)
                else:
                    if este_xml_in_memory(data):
                        new_xml = out_path + ".xml"
                        os.rename(out_path, new_xml)
                        extracted_files.append(new_xml)
                    else:
                        os.remove(out_path)

    doc.close()
    return extracted_files


########################################################################
# 3) Interfață Grafică PyQt
########################################################################

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF->XML->CSV (PyMuPDF), rânduri separate pentru taguri repetate")

        # Variabile
        self.selected_pdf_paths = []  # Liste de PDF
        self.xml_fields = set()       # Tag-uri descoperite
        self.field_mapping = {}       # {xml_tag: csv_header}
        self.config_file = "mapping_config.json"

        self.load_mapping_config()

        # UI
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout_principal = QVBoxLayout(main_widget)

        # Butoane de sus
        butoane_layout = QHBoxLayout()
        layout_principal.addLayout(butoane_layout)

        self.btn_select_pdfs = QPushButton("Selectează PDF-uri")
        self.btn_select_pdfs.clicked.connect(self.select_pdfs)
        butoane_layout.addWidget(self.btn_select_pdfs)

        self.btn_discover_fields = QPushButton("Descoperă câmpuri")
        self.btn_discover_fields.setToolTip("Extragere in-memory a fișierelor XML, parse, și adunare tag-uri unice.")
        self.btn_discover_fields.clicked.connect(self.discover_xml_fields)
        butoane_layout.addWidget(self.btn_discover_fields)

        self.btn_save_mapping = QPushButton("Salvează maparea")
        self.btn_save_mapping.clicked.connect(self.save_mapping_config)
        butoane_layout.addWidget(self.btn_save_mapping)

        self.btn_process_csv = QPushButton("Procesează -> CSV")
        self.btn_process_csv.clicked.connect(self.process_to_csv)
        butoane_layout.addWidget(self.btn_process_csv)

        self.info_label = QLabel("Niciun fișier PDF selectat.")
        layout_principal.addWidget(self.info_label)

        # Tabel (Câmp XML -> Coloană CSV)
        self.table = QTableWidget()
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["Câmp XML", "Coloană CSV"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout_principal.addWidget(self.table)

        # Meniu
        meniu = self.menuBar()
        meniu_fisier = meniu.addMenu("Fișier")
        actiune_iesire = QAction("Ieșire", self)
        actiune_iesire.triggered.connect(self.close)
        meniu_fisier.addAction(actiune_iesire)

        self.resize(900, 600)

    # ------------------------------------
    # 3.1) Selectare PDF-uri
    # ------------------------------------
    def select_pdfs(self):
        fisiere, _ = QFileDialog.getOpenFileNames(
            self,
            "Alege fișiere PDF",
            os.getcwd(),
            "Fișiere PDF (*.pdf)"
        )
        if fisiere:
            self.selected_pdf_paths = fisiere
            self.info_label.setText(f"{len(fisiere)} fișier(e) PDF selectate.")

    # ------------------------------------
    # 3.2) Descoperire Tag-uri (in memory)
    # ------------------------------------
    def discover_xml_fields(self):
        """
        Extragere in-memory => parse => colectăm tagurile unice, fără a scrie fișiere pe disk.
        Evităm duplicarea fișierelor.
        """
        if not self.selected_pdf_paths:
            QMessageBox.warning(self, "Nicio selecție", "Selectează fișiere PDF mai întâi.")
            return

        self.xml_fields.clear()
        nr_xml_total = 0
        for pdf_path in self.selected_pdf_paths:
            # Obținem set de tag-uri in memory
            fields_in_pdf = discover_xml_fields_in_memory(pdf_path)
            if fields_in_pdf:
                nr_xml_total += 1  # PDF conține măcar 1 fișier XML
            self.xml_fields.update(fields_in_pdf)

        if not self.xml_fields:
            QMessageBox.information(
                self,
                "Nu s-au găsit fișiere XML",
                "Niciun fișier XML nu a fost descoperit în PDF-urile selectate."
            )
        else:
            QMessageBox.information(
                self,
                "Gata",
                f"Descoperite {len(self.xml_fields)} tag(uri) unice din {nr_xml_total} PDF-uri cu XML."
            )

        self.populate_table()

    def populate_table(self):
        self.table.setRowCount(len(self.xml_fields))
        for row_idx, field in enumerate(sorted(self.xml_fields)):
            item_field = QTableWidgetItem(field)
            item_field.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            self.table.setItem(row_idx, 0, item_field)

            existing_csv = self.field_mapping.get(field, "")
            item_csv = QTableWidgetItem(existing_csv)
            self.table.setItem(row_idx, 1, item_csv)

    # ------------------------------------
    # 3.3) Load / Save Mapare
    # ------------------------------------
    def load_mapping_config(self):
        if os.path.isfile(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    self.field_mapping = json.load(f)
                print(f"Maparea încărcată din {self.config_file}")
            except Exception as e:
                print(f"Eroare la încărcarea mapării: {e}")

    def save_mapping_config(self):
        new_mapping = {}
        row_count = self.table.rowCount()
        for r in range(row_count):
            item_field = self.table.item(r, 0)
            item_csv = self.table.item(r, 1)
            if item_field and item_csv:
                xml_tag = item_field.text().strip()
                csv_header = item_csv.text().strip()
                if xml_tag and csv_header:
                    new_mapping[xml_tag] = csv_header

        self.field_mapping = new_mapping
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.field_mapping, f, indent=2, ensure_ascii=False)
            QMessageBox.information(self, "Mapare salvată", "Maparea a fost salvată cu succes!")
        except Exception as e:
            QMessageBox.critical(self, "Eroare", f"Eroare la salvarea mapării:\n{e}")

    # ------------------------------------
    # 3.4) Process -> CSV
    # ------------------------------------
    def process_to_csv(self):
        """
        Creează subfolder "extracted_xml/<timestamp>", extrage acolo fișiere .xml,
        apoi parsează fiecare => returnăm mai multe rânduri dacă avem repetări de tag.
        CSV final => "output_<timestamp>.csv".
        """
        if not self.selected_pdf_paths:
            QMessageBox.warning(self, "Nicio selecție", "Selectează fișiere PDF mai întâi.")
            return
        if not self.field_mapping:
            QMessageBox.warning(self, "Fără mapare", "Definește și salvează măcar o mapare de câmpuri.")
            return

        # 1) Creăm subfolder
        timestamp_str = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
        main_extracted = os.path.join(os.getcwd(), "extracted_xml")
        os.makedirs(main_extracted, exist_ok=True)

        folder_extrase = os.path.join(main_extracted, timestamp_str)
        os.makedirs(folder_extrase, exist_ok=True)

        # 2) CSV => "output_<timestamp>.csv"
        csv_path = os.path.join(os.getcwd(), f"output_{timestamp_str}.csv")

        # 3) Parcurgem PDF-urile, extragem XML pe disk, parse => multiple rows
        all_rows = []
        total_xml_found = 0
        for pdf_path in self.selected_pdf_paths:
            xml_files = extract_xml_attachments_to_disk(pdf_path, folder_extrase)
            total_xml_found += len(xml_files)
            for xfile in xml_files:
                # parse & produce multiple row dicts
                row_dicts = parse_xml_and_extract_rows(xfile, self.field_mapping)
                all_rows.extend(row_dicts)

        if not all_rows:
            QMessageBox.information(
                self,
                "Fără date",
                f"Nu s-a găsit nicio informație XML de exportat (am extras {total_xml_found} fișiere XML)."
            )
            return

        # 4) Scriem CSV
        csv_headers = list(self.field_mapping.values())
        try:
            with open(csv_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=csv_headers)
                writer.writeheader()
                for row_data in all_rows:
                    writer.writerow(row_data)

            QMessageBox.information(
                self,
                "Succes",
                f"Am extras {total_xml_found} fișier(e) XML în:\n{folder_extrase}\n\n"
                f"Am generat fișierul CSV:\n{csv_path}\n\n"
                f"Rânduri create: {len(all_rows)}"
            )
        except Exception as e:
            QMessageBox.critical(
                self,
                "Eroare",
                f"Nu am reușit să scriem fișierul CSV:\n{e}"
            )


########################################################################
# 4) MAIN
########################################################################

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
