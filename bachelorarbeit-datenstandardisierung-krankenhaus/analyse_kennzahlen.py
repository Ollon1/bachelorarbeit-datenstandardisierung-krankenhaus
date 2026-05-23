import pandas as pd
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
ANALYSIS_TABLE = BASE_DIR / "analyse_tabelle.csv"

df = pd.read_csv(ANALYSIS_TABLE, sep=";")

# 1. Durchschnittliche Auslastungsquote pro Abteilung

auslastung = (
    df.groupby("abteilung")["auslastungsquote"]
    .mean()
    .reset_index()
)

auslastung["auslastung_prozent"] = (auslastung["auslastungsquote"] * 100).round(1)

print("Durchschnittliche Auslastungsquote pro Abteilung:\n")
print(
    auslastung[["abteilung", "auslastung_prozent"]]
    .to_string(
        index=False,
        formatters={"auslastung_prozent": lambda x: f"{x:.1f} %"}
    )
)

print("\n")




# 2. Durchschnittliche Kosten pro belegtem Bett

kosten = (
    df.groupby("abteilung")["kosten_pro_belegtem_bett"]
    .mean()
    .reset_index()
)

kosten["kosten_pro_belegtem_bett"] = kosten["kosten_pro_belegtem_bett"].round(2)

print("Durchschnittliche Kosten pro belegtem Bett:\n")
print(
    kosten.to_string(
        index=False,
        formatters={"kosten_pro_belegtem_bett": lambda x: f"{x:,.2f} EUR"}
    )
)
