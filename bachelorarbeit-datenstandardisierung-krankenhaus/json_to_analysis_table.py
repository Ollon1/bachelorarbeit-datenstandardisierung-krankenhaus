"""
Tabellarische Aufbereitung der standardisierten Krankenhausdaten

Input:
  - krankenhausdaten_standardisiert.json

Output:
  - analyse_tabelle.csv

Ziel:
  Die hierarchische JSON-Zielstruktur wird in eine flache Analyse-Tabelle
  überführt. Pro Monat und Abteilung entsteht eine Tabellenzeile.

Hinweis:
  Die Kennzahlen auslastungsquote, kosten_pro_belegtem_bett und
  fte_pro_belegtem_bett werden bewusst bereits mitberechnet, damit
  die CSV-Datei direkt als Grundlage für einfache Auswertungen genutzt
  werden kann.
"""

import csv
import json
from pathlib import Path


# KONFIGURATION

BASE_DIR = Path(__file__).parent

INPUT_JSON = BASE_DIR / "krankenhausdaten_standardisiert.json"
OUTPUT_CSV = BASE_DIR / "analyse_tabelle.csv"


# HILFSFUNKTIONEN

def safe_number(value):
    """Gibt numerische Werte als float zurück; fehlende Werte werden zu 0."""
    if value is None:
        return 0.0
    return float(value)


def safe_divide(numerator, denominator):
    """Verhindert Division durch 0."""
    numerator = safe_number(numerator)
    denominator = safe_number(denominator)

    if denominator == 0:
        return 0.0

    return numerator / denominator


def round_value(value, digits=2):
    """Rundet Zahlen einheitlich."""
    return round(float(value), digits)


# JSON EINLESEN

with open(INPUT_JSON, "r", encoding="utf-8") as file:
    daten = json.load(file)


# JSON IN FLACHE TABELLENSTRUKTUR ÜBERFÜHREN

tabellenzeilen = []

for eintrag in daten:
    monat = eintrag.get("monat")
    abteilung = eintrag.get("abteilung")

    betten = eintrag.get("bettenbelegung", {})
    personalplanung = eintrag.get("personalplanung", [])
    ressourcen_kosten = eintrag.get("ressourcen_kosten", [])

    geplante_betten = safe_number(betten.get("geplante_betten"))
    tatsaechliche_betten = safe_number(betten.get("tatsaechliche_betten"))
    belegte_betten = safe_number(betten.get("belegte_betten"))
    sperrzeiten = safe_number(betten.get("sperrzeiten"))

    # Personalwerte aggregieren
    fte_arzt = 0.0
    fte_pflege = 0.0
    fte_technik = 0.0
    fte_verwaltung = 0.0

    for personal in personalplanung:
        berufsgruppe = personal.get("berufsgruppe")
        fte = safe_number(personal.get("vollzeitaequivalente_fte"))

        if berufsgruppe == "Arzt":
            fte_arzt += fte
        elif berufsgruppe == "Pflege":
            fte_pflege += fte
        elif berufsgruppe == "Technik":
            fte_technik += fte
        elif berufsgruppe == "Verwaltung":
            fte_verwaltung += fte

    gesamt_fte = fte_arzt + fte_pflege + fte_technik + fte_verwaltung

    # Kostenwerte nach Kostenart aggregieren
    personalkosten = 0.0
    materialkosten = 0.0
    energiekosten = 0.0

    for kosten in ressourcen_kosten:
        kostenart = kosten.get("kostenart")
        betrag = safe_number(kosten.get("betrag"))

        if kostenart == "Personal":
            personalkosten += betrag
        elif kostenart == "Material":
            materialkosten += betrag
        elif kostenart == "Energie":
            energiekosten += betrag

    gesamtkosten = personalkosten + materialkosten + energiekosten

    # Einfache Kennzahlen für spätere Auswertung
    auslastungsquote = safe_divide(belegte_betten, tatsaechliche_betten)
    kosten_pro_belegtem_bett = safe_divide(gesamtkosten, belegte_betten)
    fte_pro_belegtem_bett = safe_divide(gesamt_fte, belegte_betten)

    zeile = {
        "monat": monat,
        "abteilung": abteilung,
        "geplante_betten": int(geplante_betten),
        "tatsaechliche_betten": int(tatsaechliche_betten),
        "belegte_betten": int(belegte_betten),
        "sperrzeiten": int(sperrzeiten),
        "fte_arzt": round_value(fte_arzt),
        "fte_pflege": round_value(fte_pflege),
        "fte_technik": round_value(fte_technik),
        "fte_verwaltung": round_value(fte_verwaltung),
        "gesamt_fte": round_value(gesamt_fte),
        "personalkosten": round_value(personalkosten),
        "materialkosten": round_value(materialkosten),
        "energiekosten": round_value(energiekosten),
        "gesamtkosten": round_value(gesamtkosten),
        "auslastungsquote": round_value(auslastungsquote, 4),
        "kosten_pro_belegtem_bett": round_value(kosten_pro_belegtem_bett),
        "fte_pro_belegtem_bett": round_value(fte_pro_belegtem_bett, 4),
    }

    tabellenzeilen.append(zeile)


# CSV EXPORTIEREN

feldnamen = [
    "monat",
    "abteilung",
    "geplante_betten",
    "tatsaechliche_betten",
    "belegte_betten",
    "sperrzeiten",
    "fte_arzt",
    "fte_pflege",
    "fte_technik",
    "fte_verwaltung",
    "gesamt_fte",
    "personalkosten",
    "materialkosten",
    "energiekosten",
    "gesamtkosten",
    "auslastungsquote",
    "kosten_pro_belegtem_bett",
    "fte_pro_belegtem_bett",
]

with open(OUTPUT_CSV, "w", encoding="utf-8-sig", newline="") as file:
    writer = csv.DictWriter(file, fieldnames=feldnamen, delimiter=";")
    writer.writeheader()
    writer.writerows(tabellenzeilen)


# KURZE KONTROLLAUSGABE

print(f"Analyse-Tabelle erstellt: {OUTPUT_CSV}")
print(f"Anzahl Tabellenzeilen: {len(tabellenzeilen)}")
