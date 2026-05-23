"""
Transformation bereinigter Krankenhausdaten in eine JSON-basierte Zielstruktur

Input:
  - bettenbelegung_clean.xlsx
  - personalplanung_clean.xlsx
  - ressourcen_kosten_clean.xlsx

Output:
  - krankenhausdaten_standardisiert.json

Ziel:
  Die zuvor getrennten Datensätze werden über Monat und Abteilung verbunden
  und in eine einheitliche, hierarchische JSON-Struktur überführt.
"""

import json
from pathlib import Path

import pandas as pd



# KONFIGURATION

BASE_DIR = Path(__file__).parent

INPUT_BETTEN = BASE_DIR / "bettenbelegung_clean.xlsx"
INPUT_PERSONAL = BASE_DIR / "personalplanung_clean.xlsx"
INPUT_KOSTEN = BASE_DIR / "ressourcen_kosten_clean.xlsx"

OUTPUT_JSON = BASE_DIR / "krankenhausdaten_standardisiert.json"



# HILFSFUNKTIONEN

def clean_value(value):
    """
    Wandelt pandas/numpy-Werte in JSON-kompatible Python-Werte um.
    NaN-Werte werden zu None.
    """
    if pd.isna(value):
        return None

    # pandas/numpy Zahlen in normale Python-Zahlen umwandeln
    if hasattr(value, "item"):
        return value.item()

    return value


def normalize_columns(df):
    """
    Entfernt führende/nachfolgende Leerzeichen aus Spaltennamen.
    """
    df = df.copy()
    df.columns = [str(col).strip() for col in df.columns]
    return df


# DATEN EINLESEN

def load_data():
    """
    Liest die drei bereinigten Excel-Dateien ein.
    """
    betten = pd.read_excel(INPUT_BETTEN)
    personal = pd.read_excel(INPUT_PERSONAL)
    kosten = pd.read_excel(INPUT_KOSTEN)

    betten = normalize_columns(betten)
    personal = normalize_columns(personal)
    kosten = normalize_columns(kosten)

    return betten, personal, kosten


# JSON-TRANSFORMATION

def create_json_structure(betten, personal, kosten):
    """
    Erstellt die gemeinsame JSON-Struktur.

    Zentrale Verknüpfung:
      - Monat
      - Abteilung
    """

    result = []

    # Alle Kombinationen aus Monat + Abteilung aus Bettenbelegung als Basis
    # Bettenbelegung ist der aggregierte Ausgangspunkt pro Monat und Abteilung.
    betten_sorted = betten.sort_values(["Monat", "Abteilung"])

    for _, betten_row in betten_sorted.iterrows():
        monat = betten_row["Monat"]
        abteilung = betten_row["Abteilung"]

        # Passende Personalzeilen für Monat + Abteilung
        personal_rows = personal[
            (personal["Monat"] == monat) &
            (personal["Abteilung"] == abteilung)
        ]

        # Passende Ressourcen/Kosten-Zeilen für Monat + Abteilung
        kosten_rows = kosten[
            (kosten["Monat"] == monat) &
            (kosten["Abteilung"] == abteilung)
        ]

        # Bettenbelegung als Einzelobjekt
        betten_objekt = {
            "geplante_betten": clean_value(betten_row.get("Geplante_Betten")),
            "tatsaechliche_betten": clean_value(betten_row.get("Tatsaechliche_Betten")),
            "belegte_betten": clean_value(betten_row.get("Belegte_Betten")),
            "sperrzeiten": clean_value(betten_row.get("Sperrzeiten")),
        }

        # Personalplanung als Liste
        personal_liste = []

        for _, p_row in personal_rows.iterrows():
            personal_liste.append({
                "kostenstelle": clean_value(p_row.get("Kostenstelle")),
                "berufsgruppe": clean_value(p_row.get("Berufsgruppe")),
                "vollzeitaequivalente_fte": clean_value(
                    p_row.get("Vollzeitaequivalente_FTE")
                ),
                "abteilungscode": clean_value(p_row.get("Abteilungscode")),
            })

        # Ressourcen/Kosten als Liste
        kosten_liste = []

        for _, k_row in kosten_rows.iterrows():
            kosten_liste.append({
                "kostenstellennummer": clean_value(k_row.get("Kostenstellennummer")),
                "kostenstellenart": clean_value(k_row.get("Kostenstellenart")),
                "leistung": clean_value(k_row.get("Leistung")),
                "kostenart": clean_value(k_row.get("Kostenart")),
                "betrag": clean_value(k_row.get("Betrag")),
                "netto": clean_value(k_row.get("Netto")),
                "brutto": clean_value(k_row.get("Brutto")),
                "deckungsbeitrag": clean_value(k_row.get("Deckungsbeitrag")),
                "anzahl_personal": clean_value(k_row.get("Anzahl_Personal")),
            })

        # Gemeinsamer integrierter Datensatz
        eintrag = {
            "monat": clean_value(monat),
            "abteilung": clean_value(abteilung),
            "bettenbelegung": betten_objekt,
            "personalplanung": personal_liste,
            "ressourcen_kosten": kosten_liste,
        }

        result.append(eintrag)

    return result


# SPEICHERN

def save_json(data):
    """
    Speichert die JSON-Struktur als Datei.
    """
    with open(OUTPUT_JSON, "w", encoding="utf-8") as file:
        json.dump(data, file, ensure_ascii=False, indent=4)

    print(f"JSON-Datei erstellt: {OUTPUT_JSON.name}")
    print(f"Anzahl integrierter Datensätze: {len(data)}")


# MAIN

def main():
    print("=" * 60)
    print("Transformation bereinigter Krankenhausdaten in JSON")
    print("=" * 60)

    betten, personal, kosten = load_data()

    print(f"Bettenbelegung eingelesen: {len(betten)} Zeilen")
    print(f"Personalplanung eingelesen: {len(personal)} Zeilen")
    print(f"Ressourcen/Kosten eingelesen: {len(kosten)} Zeilen")

    json_data = create_json_structure(betten, personal, kosten)

    save_json(json_data)

    print("=" * 60)
    print("Transformation abgeschlossen")
    print("=" * 60)


if __name__ == "__main__":
    main()