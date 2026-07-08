"""
Dokumenten- & Rechnungsverarbeitung - kostenlose, lokale Web-App.

Firmen laden ihre Rechnungen/Belege hoch (PDF, Foto, Scan). Die App liest
die wichtigsten Felder automatisch aus, zeigt sie in einer editierbaren
Tabelle zur Kontrolle und exportiert alles als Excel-Datei.

Kein Cloud-Dienst, keine API-Kosten, keine laufenden Gebühren: Texterkennung
und Feldextraktion laufen komplett lokal (pdfplumber/PyMuPDF + Tesseract OCR
+ Regex-Heuristiken). Damit eignet es sich auch für Betriebe, die ihre
Belege aus Datenschutzgründen nicht in die Cloud schicken wollen.
"""

from __future__ import annotations

import hashlib
from dataclasses import asdict
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

from extractor import TESSERACT_AVAILABLE, extract
from parser import parse

st.set_page_config(page_title="Rechnungsverarbeitung", page_icon="📄", layout="wide")

BEISPIEL_DOKUMENTE = ["rechnung_dachdecker.pdf", "rechnung_hotel.pdf", "rechnung_restaurant.pdf"]
KONTAKT_HINWEIS = "Für deine eigenen Dokumente inkl. Anbindung an deine Buchhaltung: sprich mich an."
DEMO_ERKLAERUNG = (
    "Die Software selbst funktioniert mit beliebigen Rechnungen und Belegen jeder Branche. "
    "Nur diese öffentliche Demo ist aus Kostenschutz-Gründen auf die drei Beispielrechnungen "
    f"unten beschränkt. {KONTAKT_HINWEIS}"
)

SPALTEN_LABELS = {
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
    "branchen_hinweis": "Branchen-Zusatzinfo",
    "extraktionsmethode": "Methode",
    "hinweise": "Hinweise",
}

BRANCHEN = ["Allgemein", "Hotel", "Restaurant", "Friseur", "Handwerker", "Dachdecker", "Möbelentsorger"]

NEUE_SPALTE_PRAEFIX = "🆕 Neue Spalte: "
NICHT_UEBERNEHMEN = "— nicht übernehmen —"


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


def _beste_spalten_vorschlaege(unsere_spalten: list[str], vorhandene_spalten: list[str]) -> dict[str, str]:
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
        vorschlaege[spalte] = treffer or (NEUE_SPALTE_PRAEFIX + spalte)
    return vorschlaege


def _mit_master_zusammenfuehren(neue_zeilen: pd.DataFrame, master_df: pd.DataFrame, mapping: dict[str, str]) -> pd.DataFrame:
    umbenannt = {}
    for unsere_spalte, ziel in mapping.items():
        if ziel == NICHT_UEBERNEHMEN:
            continue
        if ziel.startswith(NEUE_SPALTE_PRAEFIX):
            umbenannt[unsere_spalte] = unsere_spalte
        else:
            umbenannt[unsere_spalte] = ziel
    neue_zeilen_umbenannt = neue_zeilen[list(umbenannt.keys())].rename(columns=umbenannt)
    return pd.concat([master_df, neue_zeilen_umbenannt], ignore_index=True, sort=False)


def main() -> None:
    demo = _demo_modus()

    st.title("📄 Dokumenten- & Rechnungsverarbeitung")
    st.caption("Automatisch die wichtigsten Daten aus Rechnungen & Belegen auslesen – branchenoffen, kostenlos, lokal.")

    if demo:
        st.info(DEMO_ERKLAERUNG, icon="👋")
        st.caption("🔒 Dateien werden nur im Arbeitsspeicher dieser Sitzung verarbeitet, nichts wird gespeichert.")

    if not TESSERACT_AVAILABLE:
        st.info(
            "OCR für gescannte/fotografierte Belege ist hier nicht aktiv. "
            "Text-PDFs funktionieren trotzdem direkt.",
            icon="ℹ️",
        )

    branche = st.selectbox("Branche (für Zusatz-Hinweise)", BRANCHEN, index=0)

    if demo:
        st.subheader("1. Beispielrechnung herunterladen")
        st.caption("Lade eine oder mehrere Beispielrechnungen herunter und zieh sie unten in das Upload-Feld.")
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

        st.subheader("2. Hochladen & Ergebnis prüfen")
        hochgeladen = st.file_uploader(
            "Beispielrechnung(en) hier hochladen",
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
                st.warning(
                    f"Diese Demo verarbeitet nur die drei bereitgestellten Beispieldateien "
                    f"(übersprungen: {', '.join(ungueltig)}). {KONTAKT_HINWEIS}"
                )
    else:
        dateien = st.file_uploader(
            "Rechnungen/Belege hochladen (PDF, PNG, JPG - auch mehrere gleichzeitig)",
            type=["pdf", "png", "jpg", "jpeg", "tif", "tiff", "bmp"],
            accept_multiple_files=True,
        )

    if not dateien:
        return

    branche_arg = "" if branche == "Allgemein" else branche

    if (
        "verarbeitete_dateien" not in st.session_state
        or st.session_state.get("verarbeitete_dateien") != [d.name for d in dateien]
    ):
        zeilen = []
        rohtexte = []
        with st.spinner(f"Verarbeite {len(dateien)} Datei(en) ..."):
            for datei in dateien:
                file_bytes = datei.getvalue()
                extraction = extract(file_bytes, datei.name)
                daten = parse(extraction.text, filename=datei.name, branche=branche_arg)
                daten.extraktionsmethode = extraction.method
                if extraction.warnings:
                    zusatz = " | ".join(extraction.warnings)
                    daten.hinweise = (daten.hinweise + " " + zusatz).strip()
                zeilen.append(asdict(daten))
                rohtexte.append(
                    {"dateiname": datei.name, "text": extraction.text, "methode": extraction.method}
                )
        st.session_state["ergebnisse"] = zeilen
        st.session_state["rohtexte"] = rohtexte
        st.session_state["verarbeitete_dateien"] = [d.name for d in dateien]

    df = pd.DataFrame(st.session_state["ergebnisse"])
    df = df.rename(columns=SPALTEN_LABELS)

    st.subheader(("3. " if demo else "") + "Extrahierte Daten (editierbar)")
    st.caption("Werte prüfen und bei Bedarf direkt in der Tabelle korrigieren.")

    bearbeitet = st.data_editor(df, num_rows="dynamic", use_container_width=True, key="editor")

    st.markdown("---")
    st.subheader("Export")

    if demo:
        excel_bytes = _to_excel(bearbeitet)
        st.download_button(
            "⬇️ Als Excel herunterladen",
            data=excel_bytes,
            file_name="rechnungsdaten_demo.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
        st.caption(f"Anhängen an eure bestehende Excel-Tabelle ist Teil der individuellen Einrichtung. {KONTAKT_HINWEIS}")
        with st.expander("Rohtext der Dokumente ansehen (zur Kontrolle)"):
            for eintrag in st.session_state.get("rohtexte", []):
                st.markdown(f"**{eintrag['dateiname']}** (Methode: {eintrag['methode']})")
                st.text(eintrag["text"][:3000] or "(kein Text erkannt)")
                st.markdown("---")
        return

    master_datei = st.file_uploader(
        "Bestehende Excel-Tabelle verknüpfen (optional) - neue Zeilen werden angehängt statt eine neue Datei zu erzeugen",
        type=["xlsx"],
        key="master_upload",
    )

    if master_datei is None:
        excel_bytes = _to_excel(bearbeitet)
        st.download_button(
            "⬇️ Als neue Excel-Datei herunterladen",
            data=excel_bytes,
            file_name="rechnungsdaten.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
    else:
        try:
            master_df = pd.read_excel(master_datei)
        except Exception as exc:
            st.error(f"Die Datei konnte nicht als Excel-Tabelle gelesen werden: {exc}")
            return

        st.caption(
            f"'{master_datei.name}' enthält {len(master_df)} bestehende Zeile(n) und "
            f"{len(master_df.columns)} Spalte(n). Ordne unten zu, in welche Spalte jedes "
            "erkannte Feld einsortiert werden soll."
        )

        vorschlaege = _beste_spalten_vorschlaege(list(bearbeitet.columns), list(master_df.columns))
        optionen = [NICHT_UEBERNEHMEN] + list(master_df.columns) + [NEUE_SPALTE_PRAEFIX + s for s in bearbeitet.columns]

        mapping = {}
        map_cols = st.columns(3)
        for i, spalte in enumerate(bearbeitet.columns):
            with map_cols[i % 3]:
                vorschlag = vorschlaege[spalte]
                index = optionen.index(vorschlag) if vorschlag in optionen else 0
                mapping[spalte] = st.selectbox(spalte, optionen, index=index, key=f"map_{spalte}")

        ziel_werte = [z for z in mapping.values() if z != NICHT_UEBERNEHMEN]
        duplikate = {z for z in ziel_werte if ziel_werte.count(z) > 1}
        if duplikate:
            st.warning(
                "Mehrere Felder sind derselben Zielspalte zugeordnet, das überschreibt sich "
                f"gegenseitig: {', '.join(duplikate)}"
            )

        zusammengefuehrt = _mit_master_zusammenfuehren(bearbeitet, master_df, mapping)
        st.caption(f"Vorschau nach Zusammenführung: {len(zusammengefuehrt)} Zeilen insgesamt")
        st.dataframe(zusammengefuehrt.tail(10), use_container_width=True)

        merged_bytes = _to_excel(zusammengefuehrt)
        st.download_button(
            "⬇️ Zusammengeführte Tabelle herunterladen",
            data=merged_bytes,
            file_name=master_datei.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
        st.caption(
            "Ersetze damit die alte Datei am selben Ort (z.B. in eurem Firmen-Ordner/OneDrive) - "
            "die neuen Rechnungen sind dann Teil der bestehenden Tabelle."
        )

    with st.expander("Rohtext der Dokumente ansehen (zur Kontrolle)"):
        for eintrag in st.session_state.get("rohtexte", []):
            st.markdown(f"**{eintrag['dateiname']}** (Methode: {eintrag['methode']})")
            st.text(eintrag["text"][:3000] or "(kein Text erkannt)")
            st.markdown("---")


if __name__ == "__main__":
    main()
