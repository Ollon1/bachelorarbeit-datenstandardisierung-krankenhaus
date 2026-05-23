"""
Bereinigungsskript: Krankenhaus-Daten
======================================
Erstellt nach manueller explorativer Analyse der 3 Excel-Dateien:
  - bettenbelegung.xlsx
  - personalplanung.xlsx
  - ressourcen_kosten.xlsx

Strategie:
  - Kategorie A (eindeutig): Automatisch bereinigen
  - Kategorie B (uneindeutig): In Quarantäne-Datei auslagern fuer Fachexperten

Output:
  - *_clean.xlsx       -> Bereinigte Daten (nur sichere Fixes)
  - quarantaene.xlsx   -> Uneindeutige Faelle fuer Expertenreview
  - changelog.csv      -> Protokoll aller Aenderungen
"""

import pandas as pd
import numpy as np
from pathlib import Path

# ============================================================
# KONFIGURATION
# ============================================================
BASE_DIR = Path(__file__).parent
INPUT_BETTEN = BASE_DIR / "bettenbelegung.xlsx"
INPUT_PERSONAL = BASE_DIR / "personalplanung.xlsx"
INPUT_KOSTEN = BASE_DIR / "ressourcen_kosten.xlsx"

OUTPUT_BETTEN = BASE_DIR / "bettenbelegung_clean.xlsx"
OUTPUT_PERSONAL = BASE_DIR / "personalplanung_clean.xlsx"
OUTPUT_KOSTEN = BASE_DIR / "ressourcen_kosten_clean.xlsx"
OUTPUT_QUARANTAENE = BASE_DIR / "quarantaene.xlsx"
OUTPUT_CHANGELOG = BASE_DIR / "changelog.csv"

# Globale Sammlungen
changelog = []
quarantaene = []


def log_change(datei, zeile, spalte, alter_wert, neuer_wert, aktion, kategorie="A"):
    """Protokolliert eine Aenderung im Changelog."""
    changelog.append({
        "Datei": datei,
        "Zeile_Original": zeile,
        "Spalte": spalte,
        "Alter_Wert": alter_wert,
        "Neuer_Wert": neuer_wert,
        "Aktion": aktion,
        "Kategorie": kategorie,
    })


def add_quarantaene(datei, zeile_nr, row_data, problem):
    """Fuegt eine Zeile zur Quarantäne hinzu."""
    entry = row_data.to_dict() if hasattr(row_data, "to_dict") else dict(row_data)
    entry["_Datei"] = datei
    entry["_Zeile_Original"] = zeile_nr
    entry["_Problem"] = problem
    quarantaene.append(entry)


# ============================================================
# 1. BETTENBELEGUNG BEREINIGEN
# ============================================================
def clean_bettenbelegung():
    print("=" * 60)
    print("1. Bereinige: bettenbelegung.xlsx")
    print("=" * 60)

    df = pd.read_excel(INPUT_BETTEN)
    datei = "bettenbelegung.xlsx"
    print(f"   Eingelesen: {len(df)} Zeilen")

    # 1a. Gross-/Kleinschreibung: Abteilung normalisieren
    abt_mapping = {
        "innere medizin": "Innere Medizin",
        "CHIRURGIE": "Chirurgie",
    }
    for alt, neu in abt_mapping.items():
        mask = df["Abteilung"] == alt
        if mask.any():
            for idx in df[mask].index:
                log_change(datei, idx, "Abteilung", alt, neu, "Schreibweise normalisiert")
            df.loc[mask, "Abteilung"] = neu
    print(f"   Schreibweisen normalisiert: {len(abt_mapping)} Regeln")

    # 1b. Duplikate entfernen
    vor = len(df)
    dup_mask = df.duplicated(keep="first")
    for idx in df[dup_mask].index:
        log_change(datei, idx, "--", "Komplette Zeile", "Entfernt",
                   "Exaktes Duplikat entfernt")
    df = df[~dup_mask].reset_index(drop=True)
    print(f"   Duplikate entfernt: {vor - len(df)}")

    # 1c. Missing Value: Belegte_Betten interpolieren
    # Sortierung sicherstellen fuer korrekte Interpolation
    df = df.sort_values(["Abteilung", "Monat"]).reset_index(drop=True)

    missing_mask = df["Belegte_Betten"].isna()
    if missing_mask.any():
        for idx in df[missing_mask].index:
            monat = df.loc[idx, "Monat"]
            abt = df.loc[idx, "Abteilung"]
            log_change(datei, idx, "Belegte_Betten", "NaN", "Interpoliert",
                       f"Zeitliche Interpolation (Nachbarmonate) fuer {monat}/{abt}")

        df["Belegte_Betten"] = (
            df.groupby("Abteilung")["Belegte_Betten"]
            .transform(lambda x: x.interpolate(method="linear"))
            .round(0)
            .astype(int)
        )
    print(f"   Missing Values interpoliert: {missing_mask.sum()}")

    # 1d. Validierung
    print("\n   --- Validierung ---")
    fehler = 0

    if df.isnull().sum().sum() > 0:
        print("   FEHLER: Noch fehlende Werte!")
        fehler += 1
    else:
        print("   OK: Keine fehlenden Werte")

    if df.duplicated().sum() > 0:
        print("   FEHLER: Noch Duplikate!")
        fehler += 1
    else:
        print("   OK: Keine Duplikate")

    if len(df) != 48:
        print(f"   WARNUNG: {len(df)} Zeilen statt erwartet 48")
        fehler += 1
    else:
        print("   OK: 48 Zeilen (12 Monate x 4 Abteilungen)")

    erwartete_abt = {"Innere Medizin", "Chirurgie", "Intensivstation", "Gynaekologie"}
    tatsaechliche_abt = set(df["Abteilung"].unique())
    if tatsaechliche_abt != erwartete_abt:
        print(f"   FEHLER: Unerwartete Abteilungen: {tatsaechliche_abt - erwartete_abt}")
        fehler += 1
    else:
        print("   OK: 4 Abteilungen korrekt")

    invalid = df[df["Belegte_Betten"] > df["Tatsaechliche_Betten"]]
    if len(invalid) > 0:
        print(f"   FEHLER: {len(invalid)}x Belegte > Tatsaechliche Betten")
        fehler += 1
    else:
        print("   OK: Belegte <= Tatsaechliche Betten")

    if fehler == 0:
        print("   ALLE CHECKS BESTANDEN")
    else:
        print(f"   {fehler} Probleme gefunden")

    return df


# ============================================================
# 2. PERSONALPLANUNG BEREINIGEN
# ============================================================
def clean_personalplanung():
    print("\n" + "=" * 60)
    print("2. Bereinige: personalplanung.xlsx")
    print("=" * 60)

    df = pd.read_excel(INPUT_PERSONAL)
    datei = "personalplanung.xlsx"
    print(f"   Eingelesen: {len(df)} Zeilen")

    # 2a. Exakte Duplikate entfernen (Zeilen 240-242)
    vor = len(df)
    dup_mask = df.duplicated(keep="first")
    for idx in df[dup_mask].index:
        log_change(datei, idx, "--", "Komplette Zeile", "Entfernt",
                   "Exaktes Duplikat entfernt")
    df = df[~dup_mask].reset_index(drop=True)
    print(f"   Exakte Duplikate entfernt: {vor - len(df)}")

    # 2b. Quarantäne-Fälle identifizieren und auslagern
    # Hinweis: Quarantäne vor den Schreibweisen-Fixes, damit die
    # Originalwerte im Quarantäne-File erhalten bleiben. Die Identifizierung
    # erfolgt anhand der Originaldaten.

    quarantaene_indices = set()

    # Fall B1: Zeile mit "ARZT" (uppercase) - Wert 1.9, koennte Verwaltung sein
    mask_arzt_upper = df["Berufsgruppe"] == "ARZT"
    for idx in df[mask_arzt_upper].index:
        add_quarantaene(datei, idx, df.loc[idx],
                        "Berufsgruppe 'ARZT' mit Wert 1.9 FTE - "
                        "koennte 'Verwaltung' sein (typischer Wert fuer Intensivstation/Verwaltung). "
                        "Gleichzeitig fehlt Verwaltung-Eintrag fuer Aug/Intensivstation.")
        quarantaene_indices.add(idx)

    # Fall B2: April/Innere Medizin/Technik doppelt (1.3 vs 1.7)
    # Nach Duplikat-Entfernung: Suche nach dem Muster
    mask_apr_inne_tech = (
        (df["Monat"] == "2024-04")
        & (df["Abteilung"].str.lower() == "innere medizin")
        & (df["Berufsgruppe"] == "Technik")
    )
    if mask_apr_inne_tech.sum() > 1:
        for idx in df[mask_apr_inne_tech].index:
            add_quarantaene(datei, idx, df.loc[idx],
                            "Doppelter Eintrag April/Innere Medizin/Technik "
                            "mit unterschiedlichen FTE-Werten (1.3 vs 1.7). "
                            "Welcher Wert ist korrekt?")
            quarantaene_indices.add(idx)

    # Fall B3: Feb/CHIRURGIE/Verwaltung (1.2) - redundant zu Zeile mit 2.0
    mask_feb_chir_extra = (
        (df["Monat"] == "2024-02")
        & (df["Abteilung"].isin(["CHIRURGIE", "Chirurgie"]))
        & (df["Berufsgruppe"] == "Verwaltung")
    )
    feb_chir_verw = df[mask_feb_chir_extra]
    if len(feb_chir_verw) > 1:
        for idx in feb_chir_verw.index:
            add_quarantaene(datei, idx, df.loc[idx],
                            "Feb/Chirurgie/Verwaltung existiert doppelt "
                            "mit unterschiedlichen Werten (2.0 vs 1.2 FTE). "
                            "Ist einer ein Duplikat oder sind beide eigenständig?")
            quarantaene_indices.add(idx)

    # Fall B4: "Chrurgie" (Tippfehler) - Abteilung klar, aber gehoert der Wert
    # wirklich zu Chirurgie/Arzt oder zu Intensivstation/Arzt (der fehlt)?
    mask_chrurgie = df["Abteilung"] == "Chrurgie"
    for idx in df[mask_chrurgie].index:
        add_quarantaene(datei, idx, df.loc[idx],
                        "Tippfehler 'Chrurgie' - Abteilung wahrscheinlich 'Chirurgie', "
                        "aber Mai/Chirurgie hat bereits einen Arzt-Eintrag (10.0 FTE). "
                        "Gleichzeitig fehlt Mai/Intensivstation/Arzt. "
                        "Gehoert dieser Wert (6.5 FTE) zu Intensivstation?")
        quarantaene_indices.add(idx)

    # Fall B5: Nov/Innere Medizin/Arzt fehlt komplett -> als "fehlend" dokumentieren
    add_quarantaene(datei, "--", pd.Series({
        "Monat": "2024-11", "Abteilung": "Innere Medizin",
        "Kostenstelle": "KST001", "Berufsgruppe": "Arzt",
        "Vollzeitaequivalente_FTE": np.nan, "Abteilungscode": "INNE",
    }), "Fehlender Eintrag: Nov/Innere Medizin/Arzt existiert nicht in den Originaldaten. "
       "Interpolation waere Okt(8.1)+Dez(7.6)/2 = 7.85, "
       "ohne Expertenbestaetigung keine Werte ableiten.")

    print(f"   Quarantäne-Fälle identifiziert: {len(quarantaene_indices)} Zeilen + 1 fehlender Eintrag")

    # Quarantäne-Zeilen aus dem Datensatz entfernen
    for idx in quarantaene_indices:
        log_change(datei, idx, "--", "--", "In Quarantäne verschoben",
                   f"Uneindeutiger Fall -> Quarantäne", kategorie="B")
    df = df.drop(index=list(quarantaene_indices)).reset_index(drop=True)

    # 2c. Schreibweisen normalisieren: Abteilung 
    abt_mapping = {
        "innere medizin": "Innere Medizin",
        "CHIRURGIE": "Chirurgie",
        "Chrurgie": "Chirurgie",  # Falls noch welche uebrig
        "Verwlatung": "Verwaltung",
    }
    for alt, neu in abt_mapping.items():
        mask = df["Abteilung"] == alt
        if mask.any():
            for idx in df[mask].index:
                log_change(datei, idx, "Abteilung", alt, neu, "Schreibweise normalisiert")
            df.loc[mask, "Abteilung"] = neu

    # 2d. Schreibweisen normalisieren: Berufsgruppe
    beruf_mapping = {
        "pflege": "Pflege",
        "ARZT": "Arzt",  # Falls noch welche uebrig (nach Quarantäne)
    }
    for alt, neu in beruf_mapping.items():
        mask = df["Berufsgruppe"] == alt
        if mask.any():
            for idx in df[mask].index:
                log_change(datei, idx, "Berufsgruppe", alt, neu, "Schreibweise normalisiert")
            df.loc[mask, "Berufsgruppe"] = neu

    print(f"   Schreibweisen normalisiert")

    # 2d2. Tippfehler "Verwlatung" erzeugt Duplikat mit bestehendem Verwaltung-Eintrag 
    # Nach Korrektur Verwlatung -> Verwaltung entsteht ein Duplikat fuer Nov/Verwaltung/Arzt.
    # Beide haben Wert 0.0, also einfach das Duplikat entfernen.
    vor2 = len(df)
    dup_mask2 = df.duplicated(keep="first")
    for idx in df[dup_mask2].index:
        log_change(datei, idx, "--", "Komplette Zeile", "Entfernt",
                   "Duplikat nach Schreibweisen-Korrektur (Verwlatung -> Verwaltung)")
    df = df[~dup_mask2].reset_index(drop=True)
    if vor2 - len(df) > 0:
        print(f"   Duplikate nach Normalisierung entfernt: {vor2 - len(df)}")

    # 2e. Fehlende Kostenstelle auffuellen 
    # Chirurgie = immer KST002
    kostenstelle_mapping = {
        "Innere Medizin": "KST001",
        "Chirurgie": "KST002",
        "Intensivstation": "KST003",
        "Gynaekologie": "KST004",
        "Verwaltung": "KST005",
    }
    mask_kst_missing = df["Kostenstelle"].isna()
    if mask_kst_missing.any():
        for idx in df[mask_kst_missing].index:
            abt = df.loc[idx, "Abteilung"]
            kst = kostenstelle_mapping.get(abt)
            if kst:
                log_change(datei, idx, "Kostenstelle", "NaN", kst,
                           f"Abgeleitet aus Abteilung '{abt}' -> {kst}")
                df.loc[idx, "Kostenstelle"] = kst
    print(f"   Fehlende Kostenstellen aufgefuellt: {mask_kst_missing.sum()}")

    # 2f. Missing FTE-Werte interpolieren 
    df = df.sort_values(["Abteilung", "Berufsgruppe", "Monat"]).reset_index(drop=True)

    missing_fte = df["Vollzeitaequivalente_FTE"].isna()
    if missing_fte.any():
        for idx in df[missing_fte].index:
            monat = df.loc[idx, "Monat"]
            abt = df.loc[idx, "Abteilung"]
            beruf = df.loc[idx, "Berufsgruppe"]
            log_change(datei, idx, "Vollzeitaequivalente_FTE", "NaN", "Interpoliert",
                       f"Zeitliche Interpolation fuer {monat}/{abt}/{beruf}")

        # Interpolation + bfill/ffill als Fallback fuer Randwerte (erster/letzter Monat)
        df["Vollzeitaequivalente_FTE"] = (
            df.groupby(["Abteilung", "Berufsgruppe"])["Vollzeitaequivalente_FTE"]
            .transform(lambda x: x.interpolate(method="linear").bfill().ffill())
        )
        # Auf 1 Dezimalstelle runden (FTE-Werte)
        df["Vollzeitaequivalente_FTE"] = df["Vollzeitaequivalente_FTE"].round(1)
    print(f"   Missing FTE interpoliert: {missing_fte.sum()}")

    # 2g. Validierung 
    print("\n   --- Validierung ---")
    fehler = 0

    remaining_missing = df.isnull().sum()
    total_missing = remaining_missing.sum()
    if total_missing > 0:
        print(f"   WARNUNG: Noch {total_missing} fehlende Werte:")
        for col, count in remaining_missing.items():
            if count > 0:
                print(f"     {col}: {count}")
        fehler += 1
    else:
        print("   OK: Keine fehlenden Werte")

    if df.duplicated().sum() > 0:
        print(f"   FEHLER: Noch {df.duplicated().sum()} Duplikate!")
        fehler += 1
    else:
        print("   OK: Keine Duplikate")

    erwartete_abt = {"Innere Medizin", "Chirurgie", "Intensivstation", "Gynaekologie", "Verwaltung"}
    tatsaechliche_abt = set(df["Abteilung"].unique())
    if tatsaechliche_abt != erwartete_abt:
        unerwartete = tatsaechliche_abt - erwartete_abt
        if unerwartete:
            print(f"   FEHLER: Unerwartete Abteilungen: {unerwartete}")
            fehler += 1
    else:
        print("   OK: 5 Abteilungen korrekt")

    erwartete_beruf = {"Arzt", "Pflege", "Verwaltung", "Technik"}
    tatsaechliche_beruf = set(df["Berufsgruppe"].unique())
    if tatsaechliche_beruf != erwartete_beruf:
        unerwartete = tatsaechliche_beruf - erwartete_beruf
        if unerwartete:
            print(f"   FEHLER: Unerwartete Berufsgruppen: {unerwartete}")
            fehler += 1
    else:
        print("   OK: 4 Berufsgruppen korrekt")

    if (df["Vollzeitaequivalente_FTE"] < 0).any():
        print("   FEHLER: Negative FTE-Werte!")
        fehler += 1
    else:
        print("   OK: Keine negativen FTE-Werte")

    if fehler == 0:
        print("   ALLE CHECKS BESTANDEN")
    else:
        print(f"   {fehler} Probleme gefunden")

    return df


# ============================================================
# 3. RESSOURCEN & KOSTEN BEREINIGEN
# ============================================================
def clean_ressourcen_kosten():
    print("\n" + "=" * 60)
    print("3. Bereinige: ressourcen_kosten.xlsx")
    print("=" * 60)

    df = pd.read_excel(INPUT_KOSTEN)
    datei = "ressourcen_kosten.xlsx"
    print(f"   Eingelesen: {len(df)} Zeilen")

    # 3a. Exakte Duplikate entfernen
    vor = len(df)
    dup_mask = df.duplicated(keep="first")
    for idx in df[dup_mask].index:
        log_change(datei, idx, "--", "Komplette Zeile", "Entfernt",
                   "Exaktes Duplikat entfernt")
    df = df[~dup_mask].reset_index(drop=True)
    print(f"   Exakte Duplikate entfernt: {vor - len(df)}")

    # 3b. Quarantäne: Deckungsbeitrag NaN (keine sichere Ableitungsregel)
    quarantaene_indices = set()

    mask_db_missing = df["Deckungsbeitrag"].isna()
    for idx in df[mask_db_missing].index:
        add_quarantaene(datei, idx, df.loc[idx],
                        f"Deckungsbeitrag = NaN fuer "
                        f"{df.loc[idx, 'Monat']}/{df.loc[idx, 'Abteilung']}/"
                        f"{df.loc[idx, 'Leistung']}/{df.loc[idx, 'Kostenart']}. "
                        f"Keine sichere Berechnungsregel ableitbar.")
        quarantaene_indices.add(idx)
        log_change(datei, idx, "Deckungsbeitrag", "NaN", "In Quarantaene verschoben",
                   "Keine sichere Ableitungsregel", kategorie="B")

    # 3c. Quarantäne: Anzahl_Personal NaN 
    # Pruefen ob der Wert aus dem gleichen Monat+Abteilung abgeleitet werden kann
    mask_personal_missing = df["Anzahl_Personal"].isna()
    for idx in df[mask_personal_missing].index:
        monat = df.loc[idx, "Monat"]
        abt = df.loc[idx, "Abteilung"]
        kostenart = df.loc[idx, "Kostenart"]

        # Suche nach gleichem Monat+Abteilung+Kostenart mit vorhandenem Wert
        same_group = df[
            (df["Monat"] == monat)
            & (df["Abteilung"] == abt)
            & (df["Kostenart"] == kostenart)
            & (df["Anzahl_Personal"].notna())
        ]

        if len(same_group) > 0:
            # Ableitung moeglich: gleicher Monat+Abteilung+Kostenart hat Wert
            wert = same_group["Anzahl_Personal"].iloc[0]
            log_change(datei, idx, "Anzahl_Personal", "NaN", wert,
                       f"Abgeleitet aus gleichem Monat/Abteilung/Kostenart ({monat}/{abt}/{kostenart})")
            df.loc[idx, "Anzahl_Personal"] = wert
        else:
            # Versuch: gleicher Monat+Abteilung (andere Kostenart)
            same_ma = df[
                (df["Monat"] == monat)
                & (df["Abteilung"] == abt)
                & (df["Anzahl_Personal"].notna())
            ]
            if len(same_ma) > 0:
                # Haeufigster Wert in dem Monat fuer die Abteilung
                wert = same_ma["Anzahl_Personal"].mode().iloc[0]
                log_change(datei, idx, "Anzahl_Personal", "NaN", wert,
                           f"Abgeleitet aus haeufgstem Wert im gleichen Monat/Abteilung ({monat}/{abt})")
                df.loc[idx, "Anzahl_Personal"] = wert
            else:
                # Nicht ableitbar -> Quarantäne
                add_quarantaene(datei, idx, df.loc[idx],
                                f"Anzahl_Personal = NaN, nicht sicher ableitbar")
                quarantaene_indices.add(idx)
                log_change(datei, idx, "Anzahl_Personal", "NaN",
                           "In Quarantäne verschoben",
                           "Nicht ableitbar", kategorie="B")

    if quarantaene_indices:
        df = df.drop(index=list(quarantaene_indices)).reset_index(drop=True)
    print(f"   Quarantäne-Fälle: {len(quarantaene_indices)} Zeilen")

    # 3d. Schreibweisen normalisieren: Abteilung
    abt_mapping = {
        "intensivstation": "Intensivstation",
    }
    for alt, neu in abt_mapping.items():
        mask = df["Abteilung"] == alt
        if mask.any():
            for idx in df[mask].index:
                log_change(datei, idx, "Abteilung", alt, neu, "Schreibweise normalisiert")
            df.loc[mask, "Abteilung"] = neu

    # 3e. Schreibweisen normalisieren: Kostenart 
    kostenart_mapping = {
        "MATERIAL": "Material",
        "Matrial": "Material",
        "energie": "Energie",
    }
    for alt, neu in kostenart_mapping.items():
        mask = df["Kostenart"] == alt
        if mask.any():
            for idx in df[mask].index:
                log_change(datei, idx, "Kostenart", alt, neu, "Schreibweise normalisiert")
            df.loc[mask, "Kostenart"] = neu

    print(f"   Schreibweisen normalisiert")

    # 3f. Fehlenden Netto-Wert berechnen (Brutto * 0.8333)
    mask_netto_missing = df["Netto"].isna()
    if mask_netto_missing.any():
        for idx in df[mask_netto_missing].index:
            brutto = df.loc[idx, "Brutto"]
            netto = round(brutto * (5 / 6), 2)  # 1/1.20 = 5/6 = 0.83333...
            log_change(datei, idx, "Netto", "NaN", netto,
                       f"Berechnet aus Brutto ({brutto}) * 5/6 (konsistente Formel ueber alle Daten)")
            df.loc[idx, "Netto"] = netto
    print(f"   Fehlende Netto-Werte berechnet: {mask_netto_missing.sum()}")

    #    3g. Redundante Spalte 'Betrag' dokumentieren 
    betrag_eq_brutto = (df["Betrag"] == df["Brutto"]).all()
    if betrag_eq_brutto:
        log_change(datei, "--", "Betrag", "== Brutto (100%)", "Spalte beibehalten, aber als redundant dokumentiert",
                   "Spalte 'Betrag' ist identisch mit 'Brutto' in allen Zeilen. "
                   "Empfehlung: Spalte entfernen oder umbenennen.", kategorie="A")
        print("   Hinweis: Spalte 'Betrag' ist 100% redundant zu 'Brutto' (dokumentiert)")

    # 3h. Validierung
    print("\n   --- Validierung ---")
    fehler = 0

    remaining_missing = df.isnull().sum()
    total_missing = remaining_missing.sum()
    if total_missing > 0:
        print(f"   WARNUNG: Noch {total_missing} fehlende Werte:")
        for col, count in remaining_missing.items():
            if count > 0:
                print(f"     {col}: {count}")
        fehler += 1
    else:
        print("   OK: Keine fehlenden Werte")

    if df.duplicated().sum() > 0:
        print(f"   FEHLER: Noch {df.duplicated().sum()} Duplikate!")
        fehler += 1
    else:
        print("   OK: Keine Duplikate")

    erwartete_abt = {"Innere Medizin", "Chirurgie", "Intensivstation", "Gynaekologie", "Verwaltung"}
    tatsaechliche_abt = set(df["Abteilung"].unique())
    if tatsaechliche_abt != erwartete_abt:
        unerwartete = tatsaechliche_abt - erwartete_abt
        if unerwartete:
            print(f"   FEHLER: Unerwartete Abteilungen: {unerwartete}")
            fehler += 1
    else:
        print("   OK: 5 Abteilungen korrekt")

    erwartete_kostenart = {"Personal", "Material", "Energie"}
    tatsaechliche_kostenart = set(df["Kostenart"].unique())
    if tatsaechliche_kostenart != erwartete_kostenart:
        unerwartete = tatsaechliche_kostenart - erwartete_kostenart
        if unerwartete:
            print(f"   FEHLER: Unerwartete Kostenarten: {unerwartete}")
            fehler += 1
    else:
        print("   OK: 3 Kostenarten korrekt")

    invalid_netto = df[df["Netto"] > df["Brutto"]].dropna(subset=["Netto", "Brutto"])
    if len(invalid_netto) > 0:
        print(f"   FEHLER: {len(invalid_netto)}x Netto > Brutto!")
        fehler += 1
    else:
        print("   OK: Netto <= Brutto")

    if (df["Betrag"] < 0).any() or (df["Netto"] < 0).any() or (df["Brutto"] < 0).any():
        print("   FEHLER: Negative Betraege!")
        fehler += 1
    else:
        print("   OK: Keine negativen Betraege")

    if fehler == 0:
        print("   ALLE CHECKS BESTANDEN")
    else:
        print(f"   {fehler} Probleme gefunden")

    return df


# ============================================================
# MAIN
# ============================================================
def main():
    print()
    print("*" * 60)
    print("  KRANKENHAUS-DATEN BEREINIGUNG")
    print("  Kategorie A: Automatisch bereinigt")
    print("  Kategorie B: In Quarantaene fuer Experten")
    print("*" * 60)

    # Bereinigung durchfuehren
    df_betten = clean_bettenbelegung()
    df_personal = clean_personalplanung()
    df_kosten = clean_ressourcen_kosten()

    # Ergebnisse speichern
    print("\n" + "=" * 60)
    print("ERGEBNISSE SPEICHERN")
    print("=" * 60)

    df_betten.to_excel(OUTPUT_BETTEN, index=False)
    print(f"   {OUTPUT_BETTEN.name}: {len(df_betten)} Zeilen")

    df_personal.to_excel(OUTPUT_PERSONAL, index=False)
    print(f"   {OUTPUT_PERSONAL.name}: {len(df_personal)} Zeilen")

    df_kosten.to_excel(OUTPUT_KOSTEN, index=False)
    print(f"   {OUTPUT_KOSTEN.name}: {len(df_kosten)} Zeilen")

    # Quarantäne speichern - aufgeteilt nach Quelldatei (separate Sheets)
    # und Schreibweisen in den Quarantäne-Daten normalisieren
    if quarantaene:
        df_q = pd.DataFrame(quarantaene)

        # Schreibweisen in Quarantäne-Daten normalisieren, damit der Experte
        # sofort erkennt, welche Abteilung gemeint ist
        abt_normalize = {
            "innere medizin": "Innere Medizin",
            "CHIRURGIE": "Chirurgie",
            "Chrurgie": "Chirurgie (Tippfehler)",
            "Verwlatung": "Verwaltung (Tippfehler)",
            "intensivstation": "Intensivstation",
        }
        if "Abteilung" in df_q.columns:
            df_q["Abteilung"] = df_q["Abteilung"].replace(abt_normalize)

        meta_cols = ["_Zeile_Original", "_Problem"]

        with pd.ExcelWriter(OUTPUT_QUARANTAENE, engine="openpyxl") as writer:
            for datei_name, gruppe in df_q.groupby("_Datei"):
                # Nur relevante Spalten (nicht-leer) fuer diese Quelldatei
                data_cols = [c for c in gruppe.columns
                             if c not in ["_Datei"] + meta_cols
                             and gruppe[c].notna().any()]
                cols = meta_cols + data_cols
                sheet_name = datei_name.replace(".xlsx", "")[:31]  # Excel max 31 chars
                gruppe[cols].to_excel(writer, sheet_name=sheet_name, index=False)

            print(f"   {OUTPUT_QUARANTAENE.name}: {len(df_q)} Faelle "
                  f"auf {df_q['_Datei'].nunique()} Sheets")
    else:
        print("   Keine Quarantaene-Faelle")

    # Changelog speichern
    if changelog:
        df_cl = pd.DataFrame(changelog)
        df_cl.to_csv(OUTPUT_CHANGELOG, index=False, sep=";", encoding="utf-8-sig")
        print(f"   {OUTPUT_CHANGELOG.name}: {len(df_cl)} Eintraege")

    # Zusammenfassung
    print("\n" + "=" * 60)
    print("ZUSAMMENFASSUNG")
    print("=" * 60)
    kat_a = sum(1 for c in changelog if c["Kategorie"] == "A")
    kat_b = sum(1 for c in changelog if c["Kategorie"] == "B")
    print(f"   Kategorie A (sicher bereinigt):     {kat_a} Aenderungen")
    print(f"   Kategorie B (Quarantaene/Experte):   {kat_b} Faelle")
    print(f"   Quarantaene-Eintraege gesamt:         {len(quarantaene)}")
    print()
    print("   Dateien erstellt:")
    print(f"     - {OUTPUT_BETTEN.name}")
    print(f"     - {OUTPUT_PERSONAL.name}")
    print(f"     - {OUTPUT_KOSTEN.name}")
    print(f"     - {OUTPUT_QUARANTAENE.name}")
    print(f"     - {OUTPUT_CHANGELOG.name}")
    print()


if __name__ == "__main__":
    main()
