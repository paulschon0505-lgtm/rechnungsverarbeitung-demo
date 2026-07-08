"""
Regelbasierte Extraktion der wichtigsten Rechnungs-/Belegfelder aus Rohtext.

Bewusst kein Cloud-KI-Modell: Rechnungen aus Deutschland/Österreich/Schweiz
folgen (unabhängig von der Branche - Hotel, Restaurant, Friseur, Handwerker,
Dachdecker, Möbelentsorger, ...) fast immer denselben Textbausteinen
("Rechnungsnummer", "Rechnungsdatum", "Gesamtbetrag", "IBAN", "USt-ID" ...).
Das macht eine kostenlose, lokale Regex-Heuristik für die Kernfelder
überraschend robust. Branchentypische Zusatzangaben (siehe BRANCHEN_KEYWORDS)
werden als Freitext-Fundstellen mitgeliefert, statt starr vorgegeben zu sein.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field


@dataclass
class Rechnungsdaten:
    dateiname: str = ""
    dokumenttyp: str = ""
    aussteller: str = ""
    rechnungsnummer: str = ""
    rechnungsdatum: str = ""
    leistungsdatum: str = ""
    faelligkeitsdatum: str = ""
    nettobetrag: str = ""
    mwst_satz: str = ""
    mwst_betrag: str = ""
    gesamtbetrag: str = ""
    iban: str = ""
    ust_id: str = ""
    branchen_hinweis: str = ""
    extraktionsmethode: str = ""
    hinweise: str = ""


AMOUNT = r"(\d{1,3}(?:[.\s]\d{3})*,\d{2})"
_MONATE = (
    r"Jan(?:uar)?|Feb(?:ruar)?|M(?:ä|ae)rz|Apr(?:il)?|Mai|Jun(?:i)?|Jul(?:i)?|"
    r"Aug(?:ust)?|Sep(?:tember)?|Okt(?:ober)?|Nov(?:ember)?|Dez(?:ember)?"
)
# Erkennt sowohl "03.07.2026" als auch "3. Juli 2026"
DATE = rf"(\d{{1,2}}[.\/]\d{{1,2}}[.\/]\d{{2,4}}|\d{{1,2}}\.?\s*(?:{_MONATE})\.?\s*\d{{2,4}})"
# Erlaubt Klammern, Zusatzworte, Prozentangaben etc. zwischen Label und Wert
# (z.B. "Zwischensumme (netto):" oder "MwSt 19%:"), aber nicht über Zeilenende hinweg.
GAP = r"[^\n]{0,25}?"


def _first_match(patterns: list[str], text: str, flags=re.IGNORECASE) -> str:
    for pattern in patterns:
        m = re.search(pattern, text, flags)
        if m:
            return m.group(1).strip()
    return ""


def _normalize_amount(raw: str) -> str:
    if not raw:
        return ""
    cleaned = raw.replace(" ", "").replace(".", "").replace(",", ".")
    try:
        return f"{float(cleaned):.2f}"
    except ValueError:
        return raw


def _guess_dokumenttyp(text: str) -> str:
    kandidaten = [
        ("Gutschrift", r"\bGutschrift\b"),
        ("Mahnung", r"\bMahnung\b"),
        ("Angebot", r"\bAngebot\b"),
        ("Lieferschein", r"\bLieferschein\b"),
        ("Rechnung", r"\bRechnung\b"),
        ("Invoice", r"\bInvoice\b"),
        ("Beleg", r"\bBeleg\b|\bKassenbon\b|\bQuittung\b"),
    ]
    for label, pattern in kandidaten:
        if re.search(pattern, text, re.IGNORECASE):
            return label
    return "Unbekannt"


def _guess_aussteller(text: str) -> str:
    zeilen = [z.strip() for z in text.splitlines() if z.strip()]
    for zeile in zeilen[:8]:
        # Firmenzeilen enthalten selten nur Ziffern/Datum und sind meist kurz
        if re.fullmatch(r"[\d\W]+", zeile):
            continue
        if len(zeile) > 60:
            continue
        if re.search(r"Rechnung|Datum|Seite|Blatt", zeile, re.IGNORECASE):
            continue
        return zeile
    return ""


BRANCHEN_KEYWORDS = {
    "Hotel": ["Zimmer", "Übernachtung", "Check-in", "Check-out", "Nächte"],
    "Restaurant": ["Tisch", "Bewirtung", "Speisen", "Getränke", "Bedienung"],
    "Friseur": ["Haarschnitt", "Färbung", "Behandlung", "Termin"],
    "Handwerker": ["Arbeitsstunden", "Std.", "Material", "Anfahrt", "Montage"],
    "Dachdecker": ["Dachfläche", "m²", "Gerüst", "Eindeckung", "Dämmung"],
    "Möbelentsorger": ["Entsorgung", "Sperrmüll", "Abholung", "Container", "m³"],
}


def _branchen_hinweis(text: str, branche: str) -> str:
    keywords = BRANCHEN_KEYWORDS.get(branche, [])
    treffer = [kw for kw in keywords if re.search(re.escape(kw), text, re.IGNORECASE)]
    return ", ".join(treffer)


def parse(text: str, filename: str = "", branche: str = "") -> Rechnungsdaten:
    daten = Rechnungsdaten(dateiname=filename)

    if not text.strip():
        daten.hinweise = "Kein Text im Dokument gefunden."
        return daten

    daten.dokumenttyp = _guess_dokumenttyp(text)
    daten.aussteller = _guess_aussteller(text)

    daten.rechnungsnummer = _first_match(
        [
            r"Rechnungs(?:nummer|nr\.?|-nr\.?)\s*[:.]?\s*([A-Za-z0-9\-\/_]+)",
            r"Rechnung\s*-?\s*(?:Nr\.?|Nummer)\s*[:.]?\s*([A-Za-z0-9\-\/_]+)",
            r"Beleg(?:nummer|nr\.?|-nr\.?)\s*[:.]?\s*([A-Za-z0-9\-\/_]+)",
            r"Invoice\s*(?:No\.?|Number|#)\s*[:.]?\s*([A-Za-z0-9\-\/_]+)",
        ],
        text,
    )

    daten.rechnungsdatum = _first_match(
        [
            rf"Rechnungsdatum\s*[:.]?\s*{DATE}",
            rf"Datum\s*[:.]?\s*{DATE}",
            rf"Invoice\s*Date\s*[:.]?\s*{DATE}",
        ],
        text,
    )

    daten.leistungsdatum = _first_match(
        [
            rf"Leistungsdatum\s*[:.]?\s*{DATE}",
            rf"Lieferdatum\s*[:.]?\s*{DATE}",
            rf"Leistungszeitraum\s*[:.]?\s*{DATE}",
        ],
        text,
    )

    daten.faelligkeitsdatum = _first_match(
        [
            rf"F(?:ä|ae)llig(?:keitsdatum)?\s*(?:bis|am)?\s*[:.]?\s*{DATE}",
            rf"Zahlbar\s*bis\s*[:.]?\s*{DATE}",
            rf"Zahlungsziel\s*[:.]?\s*{DATE}",
            rf"Due\s*Date\s*[:.]?\s*{DATE}",
        ],
        text,
    )

    daten.nettobetrag = _normalize_amount(
        _first_match(
            [
                rf"(?:Nettobetrag|Netto(?:summe)?|Zwischensumme|Subtotal){GAP}{AMOUNT}",
            ],
            text,
        )
    )

    daten.mwst_satz = _first_match(
        [
            r"(\d{1,2}(?:[.,]\d+)?)\s?%\s*(?:MwSt|USt|Umsatzsteuer|Mehrwertsteuer|VAT)",
            r"(?:MwSt|USt|Umsatzsteuer|Mehrwertsteuer|VAT)[^%\n]{0,20}?(\d{1,2}(?:[.,]\d+)?)\s?%",
        ],
        text,
    )

    daten.mwst_betrag = _normalize_amount(
        _first_match(
            [
                rf"(?:MwSt\.?|USt\.?|Umsatzsteuer|Mehrwertsteuer|VAT){GAP}{AMOUNT}",
            ],
            text,
        )
    )

    daten.gesamtbetrag = _normalize_amount(
        _first_match(
            [
                rf"(?:Gesamtbetrag|Rechnungsbetrag|Gesamtsumme|Endbetrag|Bruttobetrag|Total\s*Amount|Grand\s*Total){GAP}{AMOUNT}",
                rf"Gesamt\b{GAP}{AMOUNT}",
                rf"Total\b{GAP}{AMOUNT}",
            ],
            text,
        )
    )

    iban_match = re.search(r"\b([A-Z]{2}\d{2}(?:\s?[A-Z0-9]{4}){2,7})\b", text)
    daten.iban = iban_match.group(1).replace(" ", "") if iban_match else ""

    daten.ust_id = _first_match(
        [
            r"USt[- ]?IdNr\.?\s*[:.]?\s*([A-Z]{2}\s?\d{9})",
            r"Ust[- ]?ID\s*[:.]?\s*([A-Z]{2}\s?\d{9})",
            r"\b(DE\s?\d{9})\b",
        ],
        text,
    )
    if daten.ust_id:
        daten.ust_id = daten.ust_id.replace(" ", "")

    if branche:
        daten.branchen_hinweis = _branchen_hinweis(text, branche)

    fehlende = [
        label
        for label, wert in [
            ("Rechnungsnummer", daten.rechnungsnummer),
            ("Rechnungsdatum", daten.rechnungsdatum),
            ("Gesamtbetrag", daten.gesamtbetrag),
        ]
        if not wert
    ]
    if fehlende:
        daten.hinweise = "Bitte prüfen/ergänzen: " + ", ".join(fehlende)

    return daten
