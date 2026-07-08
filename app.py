"""
Dokumenten- & Rechnungsverarbeitung - kostenlose, lokale Web-App.

Firmen laden ihre Rechnungen/Belege hoch (PDF, Foto, Scan). Die App liest
die wichtigsten Felder automatisch aus, zeigt sie in einer editierbaren
Tabelle zur Kontrolle und exportiert alles als Excel-Datei.

Kein Cloud-Dienst, keine API-Kosten, keine laufenden Gebühren: Texterkennung
und Feldextraktion laufen komplett lokal (pdfplumber/PyMuPDF + Tesseract OCR
+ Regex-Heuristiken). Damit eignet es sich auch für Betriebe, die ihre
Belege aus Datenschutzgründen nicht in die Cloud schicken wollen.

Sprache wird über den URL-Parameter ?lang=de|en gesteuert (Standard: de),
so kann eine einbettende Seite (z.B. iframe im Portfolio) die Sprache mit
ihrem eigenen Sprachumschalter synchron halten.
"""

from __future__ import annotations

import hashlib
from dataclasses import asdict
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

from extractor import TESSERACT_AVAILABLE, extract
from parser import fehlende_pflichtfelder, parse

st.set_page_config(page_title="Rechnungsverarbeitung", page_icon="📄", layout="wide")

BEISPIEL_DOKUMENTE = ["rechnung_dachdecker.pdf", "rechnung_hotel.pdf", "rechnung_restaurant.pdf"]

SPALTEN_LABELS = {
    "de": {
        "dateiname": "Datei",
        "dokumenttyp": "Dokumenttyp",
        "aussteller": "Aussteller",
        "rechnungsnummer": "Rechnungsnr.",
        "rechnungsdatum": "Rechnungsdatum",
        "leistungsdatum": "Leistungsdatum",
        "faelligkeitsdatum": "Fälligkeitsdatum",
        "nettobetrag": "Netto (€)",
        "mwst_satz": "MwSt-Satz (%)",
        "mwst_betrag": "MwSt (€)",
        "gesamtbetrag": "Gesamtbetrag (€)",
        "iban": "IBAN",
        "ust_id": "USt-ID",
        "extraktionsmethode": "Methode",
        "hinweise": "Hinweise",
    },
    "en": {
        "dateiname": "File",
        "dokumenttyp": "Document Type",
        "aussteller": "Issuer",
        "rechnungsnummer": "Invoice No.",
        "rechnungsdatum": "Invoice Date",
        "leistungsdatum": "Service Date",
        "faelligkeitsdatum": "Due Date",
        "nettobetrag": "Net (€)",
        "mwst_satz": "VAT Rate (%)",
        "mwst_betrag": "VAT (€)",
        "gesamtbetrag": "Total (€)",
        "iban": "IBAN",
        "ust_id": "VAT ID",
        "extraktionsmethode": "Method",
        "hinweise": "Notes",
    },
}

NEUE_SPALTE_PRAEFIX = {"de": "🆕 Neue Spalte: ", "en": "🆕 New column: "}
NICHT_UEBERNEHMEN = {"de": "— nicht übernehmen —", "en": "— skip —"}

TXT = {
    "de": {
        "title": "📄 Dokumenten- & Rechnungsverarbeitung",
        "caption": "Automatisch die wichtigsten Daten aus Rechnungen & Belegen auslesen – branchenoffen, kostenlos, lokal.",
        "kontakt": "Für deine eigenen Dokumente inkl. Anbindung an deine Buchhaltung: sprich mich an.",
        "demo_banner": (
            "Die Software selbst funktioniert mit beliebigen Rechnungen und Belegen jeder Branche. "
            "Nur diese öffentliche Demo ist aus Kostenschutz-Gründen auf die drei Beispielrechnungen "
            "unten beschränkt. {kontakt}"
        ),
        "ocr_missing": "OCR für gescannte/fotografierte Belege ist hier nicht aktiv. Text-PDFs funktionieren trotzdem direkt.",
        "step1_header": "1. Beispielrechnung herunterladen",
        "step1_caption": "Lade eine oder mehrere Beispielrechnungen herunter und zieh sie unten in das Upload-Feld.",
        "step2_header": "2. Hochladen & Ergebnis prüfen",
        "upload_demo_label": "Beispielrechnung(en) hier hochladen",
        "upload_full_label": "Rechnungen/Belege hochladen (PDF, PNG, JPG - auch mehrere gleichzeitig)",
        "invalid_files_warning": "Diese Demo verarbeitet nur die drei bereitgestellten Beispieldateien (übersprungen: {names}). {kontakt}",
        "step3_prefix": "3. ",
        "results_header": "Extrahierte Daten (editierbar)",
        "results_caption": "Werte prüfen und bei Bedarf direkt in der Tabelle korrigieren.",
        "export_header": "Export",
        "download_excel": "⬇️ Als Excel herunterladen",
        "download_excel_full": "⬇️ Als neue Excel-Datei herunterladen",
        "export_caption_demo": "Anhängen an eure bestehende Excel-Tabelle ist Teil der individuellen Einrichtung. {kontakt}",
        "raw_text_expander": "Rohtext der Dokumente ansehen (zur Kontrolle)",
        "no_text_found": "(kein Text erkannt)",
        "method_label": "Methode",
        "processing_spinner": "Verarbeite {n} Datei(en) ...",
        "missing_fields_prefix": "Bitte prüfen/ergänzen: ",
        "upload_master_label": "Bestehende Excel-Tabelle verknüpfen (optional) - neue Zeilen werden angehängt statt eine neue Datei zu erzeugen",
        "master_read_error": "Die Datei konnte nicht als Excel-Tabelle gelesen werden: {exc}",
        "master_info": "'{name}' enthält {rows} bestehende Zeile(n) und {cols} Spalte(n). Ordne unten zu, in welche Spalte jedes erkannte Feld einsortiert werden soll.",
        "duplicate_mapping_warning": "Mehrere Felder sind derselben Zielspalte zugeordnet, das überschreibt sich gegenseitig: {names}",
        "merge_preview_caption": "Vorschau nach Zusammenführung: {n} Zeilen insgesamt",
        "download_merged": "⬇️ Zusammengeführte Tabelle herunterladen",
        "merge_replace_hint": "Ersetze damit die alte Datei am selben Ort (z.B. in eurem Firmen-Ordner/OneDrive) - die neuen Rechnungen sind dann Teil der bestehenden Tabelle.",
    },
    "en": {
        "title": "📄 Document & Invoice Processing",
        "caption": "Automatically extract the key data from invoices & receipts – works across industries, free, local.",
        "kontakt": "For your own documents, incl. integration with your accounting: get in touch with me.",
        "demo_banner": (
            "The software itself works with any invoice or receipt from any industry. Only this public "
            "demo is limited to the three sample invoices below, to keep it free for everyone to try. {kontakt}"
        ),
        "ocr_missing": "OCR for scanned/photographed receipts isn't active here. Text-based PDFs still work directly.",
        "step1_header": "1. Download a sample invoice",
        "step1_caption": "Download one or more sample invoices, then drag them into the upload field below.",
        "step2_header": "2. Upload & review the result",
        "upload_demo_label": "Upload the sample invoice(s) here",
        "upload_full_label": "Upload invoices/receipts (PDF, PNG, JPG - multiple files supported)",
        "invalid_files_warning": "This demo only processes the three provided sample files (skipped: {names}). {kontakt}",
        "step3_prefix": "3. ",
        "results_header": "Extracted Data (editable)",
        "results_caption": "Review the values and correct them directly in the table if needed.",
        "export_header": "Export",
        "download_excel": "⬇️ Download as Excel",
        "download_excel_full": "⬇️ Download as new Excel file",
        "export_caption_demo": "Appending to your existing company Excel sheet is part of the custom setup. {kontakt}",
        "raw_text_expander": "View raw extracted text (for verification)",
        "no_text_found": "(no text detected)",
        "method_label": "Method",
        "processing_spinner": "Processing {n} file(s) ...",
        "missing_fields_prefix": "Please check/complete: ",
        "upload_master_label": "Link an existing Excel sheet (optional) - new rows get appended instead of creating a new file",
        "master_read_error": "Couldn't read this file as an Excel sheet: {exc}",
        "master_info": "'{name}' has {rows} existing row(s) and {cols} column(s). Below, map which column each extracted field should go into.",
        "duplicate_mapping_warning": "Multiple fields are mapped to the same target column, which will overwrite each other: {names}",
        "merge_preview_caption": "Preview after merging: {n} rows total",
        "download_merged": "⬇️ Download merged table",
        "merge_replace_hint": "Replace the old file in the same location (e.g. your company folder/OneDrive) with this one - the new invoices are now part of the existing table.",
    },
}


def _get_lang() -> str:
    try:
        val = st.query_params.get("lang", "de")
    except Exception:
        val = "de"
    return val if val in ("de", "en") else "de"


def _demo_modus() -> bool:
    try:
        return str(st.secrets.get("DEMO_MODE", "false")).lower() == "true"
    except Exception:
        return False


def _beispiel_hashes() -> dict[str, str]:
    """sha256 -> Dateiname der mitgelieferten Beispieldokumente."""
    hashes = {}
    for name in BEISPIEL_DOKUMENTE:
        daten = (Path(__file__).parent / name).read_bytes()
        hashes[hashlib.sha256(daten).hexdigest()] = name
    return hashes


def _to_excel(df: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Rechnungen")
        sheet = writer.sheets["Rechnungen"]
        for i, col in enumerate(df.columns, start=1):
            content_width = df[col].fillna("").astype(str).map(len).max() if len(df) else 0
            width = max(12, min(40, max(content_width, len(col)) + 2))
            sheet.column_dimensions[sheet.cell(row=1, column=i).column_letter].width = width
    return buffer.getvalue()


def _normalize_spaltenname(name: str) -> str:
    ersatz = str(name).lower()
    for a, b in [("ä", "a"), ("ö", "o"), ("ü", "u"), ("ß", "ss")]:
        ersatz = ersatz.replace(a, b)
    return "".join(ch for ch in ersatz if ch.isalnum())


def _beste_spalten_vorschlaege(unsere_spalten: list[str], vorhandene_spalten: list[str], lang: str) -> dict[str, str]:
    """Rät für jede unserer Spalten die wahrscheinlich passende Spalte der Firmen-Tabelle."""
    normalisiert_vorhanden = {_normalize_spaltenname(sp): sp for sp in vorhandene_spalten}
    vorschlaege = {}
    for spalte in unsere_spalten:
        norm = _normalize_spaltenname(spalte)
        treffer = normalisiert_vorhanden.get(norm)
        if not treffer:
            for norm_vh, original in normalisiert_vorhanden.items():
                if norm in norm_vh or norm_vh in norm:
                    treffer = original
                    break
        vorschlaege[spalte] = treffer or (NEUE_SPALTE_PRAEFIX[lang] + spalte)
    return vorschlaege


def _mit_master_zusammenfuehren(neue_zeilen: pd.DataFrame, master_df: pd.DataFrame, mapping: dict[str, str], lang: str) -> pd.DataFrame:
    umbenannt = {}
    for unsere_spalte, ziel in mapping.items():
        if ziel == NICHT_UEBERNEHMEN[lang]:
            continue
        if ziel.startswith(NEUE_SPALTE_PRAEFIX[lang]):
            umbenannt[unsere_spalte] = unsere_spalte
        else:
            umbenannt[unsere_spalte] = ziel
    neue_zeilen_umbenannt = neue_zeilen[list(umbenannt.keys())].rename(columns=umbenannt)
    return pd.concat([master_df, neue_zeilen_umbenannt], ignore_index=True, sort=False)


def main() -> None:
    lang = _get_lang()
    t = TXT[lang]
    demo = _demo_modus()

    st.title(t["title"])
    st.caption(t["caption"])

    if demo:
        st.info(t["demo_banner"].format(kontakt=t["kontakt"]), icon="👋")

    if not TESSERACT_AVAILABLE:
        st.info(t["ocr_missing"], icon="ℹ️")

    if demo:
        st.subheader(t["step1_header"])
        st.caption(t["step1_caption"])
        beispiel_cols = st.columns(len(BEISPIEL_DOKUMENTE))
        for col, name in zip(beispiel_cols, BEISPIEL_DOKUMENTE):
            with col:
                st.download_button(
                    f"⬇️ {name}",
                    data=(Path(__file__).parent / name).read_bytes(),
                    file_name=name,
                    key=f"dl_{name}",
                    use_container_width=True,
                )

        st.subheader(t["step2_header"])
        hochgeladen = st.file_uploader(
            t["upload_demo_label"],
            type=["pdf", "png", "jpg", "jpeg", "tif", "tiff", "bmp"],
            accept_multiple_files=True,
        )
        dateien = []
        if hochgeladen:
            bekannte_hashes = _beispiel_hashes()
            ungueltig = []
            for d in hochgeladen:
                if hashlib.sha256(d.getvalue()).hexdigest() in bekannte_hashes:
                    dateien.append(d)
                else:
                    ungueltig.append(d.name)
            if ungueltig:
                st.warning(t["invalid_files_warning"].format(names=", ".join(ungueltig), kontakt=t["kontakt"]))
    else:
        dateien = st.file_uploader(
            t["upload_full_label"],
            type=["pdf", "png", "jpg", "jpeg", "tif", "tiff", "bmp"],
            accept_multiple_files=True,
        )

    if not dateien:
        return

    if (
        "verarbeitete_dateien" not in st.session_state
        or st.session_state.get("verarbeitete_dateien") != [d.name for d in dateien]
    ):
        zeilen = []
        rohtexte = []
        with st.spinner(t["processing_spinner"].format(n=len(dateien))):
            for datei in dateien:
                file_bytes = datei.getvalue()
                extraction = extract(file_bytes, datei.name)
                daten = parse(extraction.text, filename=datei.name)
                daten.extraktionsmethode = extraction.method
                hinweise_teile = []
                fehlende = fehlende_pflichtfelder(daten)
                if fehlende:
                    labels = [SPALTEN_LABELS[lang][f] for f in fehlende]
                    hinweise_teile.append(t["missing_fields_prefix"] + ", ".join(labels))
                if extraction.warnings:
                    hinweise_teile.append(" | ".join(extraction.warnings))
                daten.hinweise = " ".join(hinweise_teile).strip()
                zeilen.append(asdict(daten))
                rohtexte.append(
                    {"dateiname": datei.name, "text": extraction.text, "methode": extraction.method}
                )
        st.session_state["ergebnisse"] = zeilen
        st.session_state["rohtexte"] = rohtexte
        st.session_state["verarbeitete_dateien"] = [d.name for d in dateien]

    df = pd.DataFrame(st.session_state["ergebnisse"])
    df = df.drop(columns=["branchen_hinweis"], errors="ignore")
    df = df.rename(columns=SPALTEN_LABELS[lang])

    st.subheader((t["step3_prefix"] if demo else "") + t["results_header"])
    st.caption(t["results_caption"])

    bearbeitet = st.data_editor(df, num_rows="dynamic", use_container_width=True, key="editor")

    st.markdown("---")
    st.subheader(t["export_header"])

    if demo:
        excel_bytes = _to_excel(bearbeitet)
        st.download_button(
            t["download_excel"],
            data=excel_bytes,
            file_name="rechnungsdaten_demo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
        st.caption(t["export_caption_demo"].format(kontakt=t["kontakt"]))
        with st.expander(t["raw_text_expander"]):
            for eintrag in st.session_state.get("rohtexte", []):
                st.markdown(f"**{eintrag['dateiname']}** ({t['method_label']}: {eintrag['methode']})")
                st.text(eintrag["text"][:3000] or t["no_text_found"])
                st.markdown("---")
        return

    master_datei = st.file_uploader(t["upload_master_label"], type=["xlsx"], key="master_upload")

    if master_datei is None:
        excel_bytes = _to_excel(bearbeitet)
        st.download_button(
            t["download_excel_full"],
            data=excel_bytes,
            file_name="rechnungsdaten.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
    else:
        try:
            master_df = pd.read_excel(master_datei)
        except Exception as exc:
            st.error(t["master_read_error"].format(exc=exc))
            return

        st.caption(t["master_info"].format(name=master_datei.name, rows=len(master_df), cols=len(master_df.columns)))

        vorschlaege = _beste_spalten_vorschlaege(list(bearbeitet.columns), list(master_df.columns), lang)
        optionen = [NICHT_UEBERNEHMEN[lang]] + list(master_df.columns) + [NEUE_SPALTE_PRAEFIX[lang] + s for s in bearbeitet.columns]

        mapping = {}
        map_cols = st.columns(3)
        for i, spalte in enumerate(bearbeitet.columns):
            with map_cols[i % 3]:
                vorschlag = vorschlaege[spalte]
                index = optionen.index(vorschlag) if vorschlag in optionen else 0
                mapping[spalte] = st.selectbox(spalte, optionen, index=index, key=f"map_{spalte}")

        ziel_werte = [z for z in mapping.values() if z != NICHT_UEBERNEHMEN[lang]]
        duplikate = {z for z in ziel_werte if ziel_werte.count(z) > 1}
        if duplikate:
            st.warning(t["duplicate_mapping_warning"].format(names=", ".join(duplikate)))

        zusammengefuehrt = _mit_master_zusammenfuehren(bearbeitet, master_df, mapping, lang)
        st.caption(t["merge_preview_caption"].format(n=len(zusammengefuehrt)))
        st.dataframe(zusammengefuehrt.tail(10), use_container_width=True)

        merged_bytes = _to_excel(zusammengefuehrt)
        st.download_button(
            t["download_merged"],
            data=merged_bytes,
            file_name=master_datei.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
        st.caption(t["merge_replace_hint"])

    with st.expander(t["raw_text_expander"]):
        for eintrag in st.session_state.get("rohtexte", []):
            st.markdown(f"**{eintrag['dateiname']}** ({t['method_label']}: {eintrag['methode']})")
            st.text(eintrag["text"][:3000] or t["no_text_found"])
            st.markdown("---")


if __name__ == "__main__":
    main()
