# Dokumenten- & Rechnungsverarbeitung

Web-App, die Rechnungen/Belege (PDF, Foto, Scan) automatisch ausliest und die
wichtigsten Daten in eine editierbare Tabelle bzw. Excel-Datei überträgt.
Branchenoffen einsetzbar: Hotels, Restaurants, Friseure, Handwerker,
Dachdecker, Möbelentsorger und mehr.

**Kostenmodell: 0 €.** Es wird kein Cloud-KI-Dienst verwendet. Texterkennung
läuft lokal über `pdfplumber`/`PyMuPDF` (für digitale PDFs) und optional
`Tesseract OCR` (für Scans/Fotos). Die Felderkennung basiert auf
Regex-Heuristiken, die typische deutsche Rechnungsformulierungen erkennen
("Rechnungsnummer", "Gesamtbetrag", "IBAN", "USt-IdNr." ...). Kein Dokument
verlässt den eigenen Rechner – das ist gerade für Firmen ein Argument, die
ihre Belege nicht in die Cloud schicken wollen.

## Extrahierte Felder

Dokumenttyp, Aussteller, Rechnungsnummer, Rechnungsdatum, Leistungsdatum,
Fälligkeitsdatum, Nettobetrag, MwSt-Satz, MwSt-Betrag, Gesamtbetrag, IBAN,
USt-IdNr. – plus optionale Branchen-Stichwörter (z. B. "Übernachtung" bei
Hotels, "Arbeitsstunden" bei Handwerkern), je nach gewählter Branche in der
Sidebar.

Alle Werte lassen sich vor dem Export direkt in der Tabelle korrigieren –
die automatische Erkennung ist ein Vorschlag, keine Blackbox.

## Setup

```bash
python -m venv venv
venv\Scripts\activate          # Windows
source venv/bin/activate       # macOS/Linux
pip install -r requirements.txt
```

### Optional: OCR für Scans/Fotos aktivieren

Text-PDFs (z. B. Rechnungen aus Buchhaltungssoftware) funktionieren direkt
nach dem obigen Setup. Für eingescannte oder fotografierte Belege wird
zusätzlich die kostenlose Tesseract-OCR-Engine benötigt:

- **Windows:** Installer von https://github.com/UB-Mannheim/tesseract/wiki
  herunterladen, danach den Installationsordner (enthält `tesseract.exe`)
  zum PATH hinzufügen.
- **macOS:** `brew install tesseract tesseract-lang`
- **Linux:** `sudo apt install tesseract-ocr tesseract-ocr-deu`

Ohne Tesseract funktioniert die App weiterhin für alle text-basierten PDFs –
es erscheint nur ein Hinweis, dass der OCR-Fallback nicht aktiv ist.

## Starten

```bash
streamlit run app.py
```

Die App öffnet sich unter http://localhost:8501. Im Ordner liegen drei
strukturell unterschiedliche Beispielrechnungen zum Ausprobieren
(`rechnung_dachdecker.pdf`, `rechnung_hotel.pdf`, `rechnung_restaurant.pdf`)
sowie `beispiel_master_tabelle.xlsx` zum Testen der "bestehende Tabelle
verknüpfen"-Funktion.

### Neue Rechnungen an eine bestehende Excel-Tabelle anhängen

Im Export-Bereich kann optional eine bereits vorhandene Firmen-Excel-Tabelle
hochgeladen werden. Die App schlägt für jedes erkannte Feld automatisch die
passende Zielspalte vor (per Namensabgleich), lässt sich aber pro Feld frei
umstellen - inklusive "als neue Spalte anlegen", falls die Firmentabelle das
Feld noch nicht kennt. Bestehende Zeilen und Spalten bleiben dabei erhalten;
heruntergeladen wird die zusammengeführte Datei, die dann die alte ersetzt.

## Projektstruktur

- `extractor.py` – Text-/OCR-Extraktion aus PDF/Bild
- `parser.py` – Regex-Heuristiken zur Felderkennung + Branchen-Stichwörter
- `app.py` – Streamlit-Oberfläche (Upload, editierbare Tabelle, Excel-Export)

## Grenzen & Weiterentwicklung

- Die Regex-Heuristiken decken die gängigsten deutschen Rechnungsformate ab,
  aber nicht jedes exotische Layout – deshalb ist die Ergebnistabelle
  bewusst editierbar.
- Für höhere Trefferquote bei sehr unterschiedlichen/internationalen
  Layouts ließe sich optional ein KI-Modell mit Dokumentenverständnis
  (z. B. über die Anthropic API) als zweite Erkennungsstufe ergänzen -
  das würde dann aber laufende Kosten pro Dokument verursachen.
- Aktuell wird nur Excel als Export angeboten; CSV/JSON oder das Befüllen
  einer firmeneigenen Word-/PDF-Vorlage lassen sich mit `pandas` bzw.
  `python-docx` leicht ergänzen, sobald eine konkrete Firma ein bestimmtes
  Zielformat braucht.
