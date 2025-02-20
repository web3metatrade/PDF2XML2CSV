# PDF2XML2CSV - Convertor PDF -> XML -> CSV

Acest proiect reprezintă un utilitar GUI (aplicație desktop) scris în **Python** (utilizând **PyQt5** și **PyMuPDF**) care extrage și convertește fișiere XML atașate în PDF-urile generate, de exemplu, de SPV Trezorerie (sau alte sisteme ce atașează XML-uri în PDF). Aplicația permite:

- Selectarea mai multor PDF-uri.
- Descoperirea automată (in-memory) a tag-urilor din fișierele XML atașate.
- Configurarea unei mapări între tag-urile XML și coloanele dorite în fișierul CSV.
- Generarea fișierului CSV final cu datele extrase.

---

## 1. Cum descarci codul?

1. Intră pe [repo-ul PDF2XML2CSV](https://github.com/web3metatrade/PDF2XML2CSV).  
2. Fă click pe butonul verde **Code** → **Download ZIP** sau clonează cu git:  
   `git clone https://github.com/web3metatrade/PDF2XML2CSV.git`  
3. După descărcare/clonare, vei avea folderul `PDF2XML2CSV` cu fișierele de proiect.

---

## 2. Cum rulezi aplicația (modul sursă Python)?

1. Asigură-te că ai instalat Python 3.7+.
2. Din interiorul folderului descărcat, instalează dependențele:  
   `pip install PyQt5 PyMuPDF lxml`
3. Rulează aplicația:  
   `python main.py`
4. Se va deschide o fereastră GUI, de unde poți:
   - Selecta fișiere PDF cu atașamente XML.
   - Apăsa „Descoperă câmpuri” pentru a vedea tag-urile XML unice.
   - Configura maparea „Tag XML” → „Coloană CSV”.
   - Salva maparea.
   - Genera fișierul CSV final.

---

## 3. Cum compilezi într-un fișier executabil (Windows)?

Dacă vrei să rulezi aplicația fără a necesita instalarea Python pe calculatorul țintă, poți crea un `.exe` cu [PyInstaller](https://pypi.org/project/PyInstaller/):

1. Instalează PyInstaller (dacă nu ai făcut-o deja):  
   `pip install pyinstaller`
2. Din folderul proiectului, execută:  
   `pyinstaller --noconfirm --onefile --windowed main.py`  
   - `--onefile` creează un singur fișier `.exe`.
   - `--windowed` împiedică deschiderea unei ferestre de consolă.
3. După rulare, vei găsi executabilul generat în folderul `dist/` (ex: `dist\main.exe`).
4. Poți distribui acest `.exe` pe un sistem Windows fără ca Python să fie instalat.

---

## 4. Utilizare pas cu pas în aplicație

1. **Selectează PDF-uri**  
   Alege unul sau mai multe fișiere PDF care conțin atașamente XML (document-level sau adnotări de tip `FileAttachment`).

2. **Descoperă câmpuri**  
   Aplicația extrage atașamentele în memorie, parsează fișierele XML și colectează toate tag-urile într-o listă unică.

3. **Mapare „Tag XML” → „Coloană CSV”**  
   În tabelul afișat, pentru fiecare tag XML, introdu numele coloanei dorite în viitorul fișier CSV.

4. **Salvează maparea**  
   Aplicația va scrie fișierul `mapping_config.json`, cu asocierea `{xml_tag: csv_header}`.

5. **Procesează → CSV**  
   - Creează un subfolder `extracted_xml/<timestamp>` unde salvează fișierele XML (pentru arhivare/inspecție).  
   - Generează fișierul `output_<timestamp>.csv` în același folder unde se află `main.py`.  
   - Dacă un tag apare de mai multe ori, se creează rânduri multiple, duplicând valorile pentru restul tag-urilor.

6. **Rezultat**  
   Vei avea fișierele XML extrase pe disk și fișierul CSV gata de import în alte aplicații.

---

## 5. Structură proiect

.
├── main.py               # Codul principal al aplicației
├── mapping_config.json   # Fișier JSON cu maparea (generat în timpul utilizării)
├── README.md             # Acest fișier de documentație
└── requirements.txt      # (Opțional) Listă de pachete necesare


- `extracted_xml` *(generat la rulare)* conține subfoldere `<timestamp>` cu fișierele XML extrase.
- `output_<timestamp>.csv` *(generat la rulare)* este fișierul CSV final.

---

## 6. Observații importante

1. **Suport PDF**  
   Aplicația extrage doar fișierele XML atașate în PDF. Dacă PDF-ul nu are atașamente XML, nu se va genera nimic.

2. **Nume fișiere**  
   Aplicația înlocuiește caracterele invalide ( `\ / : * ? " < > | \r \n` ) cu `_` la salvarea fișierelor XML.

3. **Performanță**  
   Pentru multe PDF-uri sau fișiere XML voluminoase, parsearea poate dura; eventual, se pot face optimizări suplimentare.

4. **Tag-uri repetitive**  
   Dacă un tag apare de mai multe ori într-un fișier XML, se vor crea mai multe rânduri în CSV, duplicând valorile pentru celelalte tag-uri.

---

## 7. Licență

[MIT License](https://opensource.org/licenses/MIT


---

## 8. Contribuții

- PR-urile și issue-urile sunt binevenite.
- Dacă întâmpini probleme, deschide un *issue* pe [GitHub](https://github.com/web3metatrade/PDF2XML2CSV).

Spor la mapare si import automat!

