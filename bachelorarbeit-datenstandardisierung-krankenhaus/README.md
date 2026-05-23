# Bachelorarbeit: Datenstandardisierung im Krankenhaus

Dieses Repository enthält die Skripte, Rohdaten, bereinigten Daten und Ergebnisdateien für einen Prototyp zur Datenbereinigung und Standardisierung synthetischer Krankenhausdaten.

## Inhalt

- `clean_krankenhaus_data.py`: bereinigt die Excel-Ausgangsdaten und erstellt bereinigte Dateien sowie ein Changelog.
- `transform_to_json.py`: überführt die bereinigten Daten in eine standardisierte JSON-Struktur.
- `json_to_analysis_table.py`: erzeugt aus der JSON-Datei eine flache Analyse-Tabelle.
- `analyse_kennzahlen.py`: berechnet einfache Kennzahlen aus der Analyse-Tabelle.
- `*.xlsx`, `*.csv`, `*.json`: Eingangs-, Zwischen- und Ergebnisdaten des Prototyps.

## Nutzung

Abhängigkeiten installieren:

```powershell
pip install -r requirements.txt
```

Skripte in Reihenfolge ausführen:

```powershell
python clean_krankenhaus_data.py
python transform_to_json.py
python json_to_analysis_table.py
python analyse_kennzahlen.py
```

## Hinweis

Die enthaltenen Daten sind synthetisch erzeugt und enthalten keine realen personenbezogenen Krankenhausdaten.
