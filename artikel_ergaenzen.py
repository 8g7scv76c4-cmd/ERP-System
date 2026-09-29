"""Ergänzt die OrgaCalc-Artikelliste (CSV aus dem Zeichenprogramm) um die
Spalten Länge, Breite, Dicke und Verschnitt.

Aufruf:
    python artikel_ergaenzen.py <Artikel.csv> [--regeln Verschnitt_Regeln.xlsx]

Ohne --regeln wird zusätzlich eine Vorlage "Verschnitt_Regeln.xlsx" erzeugt:
eine Zeile pro Hauptgruppe/Gruppe, in die man den Verschnitt einträgt.
Optional kann in der Spalte "Stichwort" ein Text stehen; die Regel gilt dann
nur für Artikel, deren RWDM-Nr., Beschreibung oder Katalogpfad ihn enthält
(z. B. Gruppe "Blum Legrabox" + Stichwort "Zargen"). Regeln mit Stichwort
haben Vorrang vor Regeln ohne.

Die Ausgabe-CSV hat dasselbe Format wie die Eingabe (Semikolon, Windows-1252,
CRLF), damit OrgaCalc sie wie gewohnt einlesen kann. Alle Werte bleiben Text,
Bestellnummern werden also nicht verändert.
"""

import argparse
import csv
from pathlib import Path

import pandas as pd

ENCODING = "cp1252"
NEUE_SPALTEN = ["Länge", "Breite", "Dicke", "Verschnitt"]
REGEL_SPALTEN = ["Hauptgruppe Name", "Gruppe Name", "Stichwort", "Verschnitt"]


def lese_artikel(pfad):
    return pd.read_csv(pfad, sep=";", encoding=ENCODING, dtype=str, keep_default_na=False)


def schreibe_regelvorlage(df, pfad):
    aktiv = df[df["Arch."] == "0"]
    uebersicht = (
        aktiv.groupby(["Hauptgruppe Name", "Gruppe Name"])
        .agg(
            Anzahl=("RWDM-Nr.", "size"),
            Beispiele=("Beschreibung", lambda s: " | ".join(s.head(3))),
        )
        .reset_index()
    )
    summe = uebersicht.groupby("Hauptgruppe Name")["Anzahl"].transform("sum")
    uebersicht = uebersicht.assign(_s=summe).sort_values(
        ["_s", "Hauptgruppe Name", "Anzahl"], ascending=[False, True, False]
    ).drop(columns="_s")
    uebersicht.insert(2, "Stichwort", "")
    uebersicht.insert(3, "Verschnitt", "")

    with pd.ExcelWriter(pfad) as writer:
        uebersicht.to_excel(writer, sheet_name="Regeln", index=False)
        ws = writer.book["Regeln"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for spalte, breite in zip("ABCDEF", (22, 28, 16, 40, 9, 100)):
            ws.column_dimensions[spalte].width = breite


def lese_regeln(pfad):
    regeln = pd.read_excel(pfad, dtype=str).fillna("")
    fehlend = set(REGEL_SPALTEN) - set(regeln.columns)
    if fehlend:
        raise SystemExit(f"In {pfad} fehlen die Spalten: {', '.join(sorted(fehlend))}")
    regeln = regeln[regeln["Verschnitt"].str.strip() != ""]
    # Regeln mit Stichwort zuerst, damit sie Vorrang haben
    return regeln.sort_values("Stichwort", key=lambda s: s == "")


def verschnitt_fuer(artikel, regeln):
    text = " ".join(
        (artikel["RWDM-Nr."], artikel["Beschreibung"], artikel["Pfad zum Katalog"])
    ).lower()
    for _, r in regeln.iterrows():
        if r["Hauptgruppe Name"] and r["Hauptgruppe Name"] != artikel["Hauptgruppe Name"]:
            continue
        if r["Gruppe Name"] and r["Gruppe Name"] != artikel["Gruppe Name"]:
            continue
        if r["Stichwort"] and r["Stichwort"].lower() not in text:
            continue
        return r["Verschnitt"].strip()
    return ""


def main():
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("csv", type=Path, help="Artikelliste (CSV mit Semikolon)")
    parser.add_argument("--regeln", type=Path, help="ausgefüllte Verschnitt_Regeln.xlsx")
    parser.add_argument("--ausgabe", type=Path, help="Ziel-CSV (Standard: <Name>_ergaenzt.csv)")
    args = parser.parse_args()

    df = lese_artikel(args.csv)
    for spalte in NEUE_SPALTEN:
        if spalte not in df.columns:
            df[spalte] = ""

    if args.regeln:
        regeln = lese_regeln(args.regeln)
        df["Verschnitt"] = df.apply(verschnitt_fuer, axis=1, regeln=regeln)
        ohne = (df["Verschnitt"] == "") & (df["Arch."] == "0")
        print(f"Verschnitt gesetzt: {(df['Verschnitt'] != '').sum()}, aktiv ohne Verschnitt: {ohne.sum()}")
    else:
        vorlage = args.csv.with_name("Verschnitt_Regeln.xlsx")
        schreibe_regelvorlage(df, vorlage)
        print(f"Regelvorlage geschrieben: {vorlage}")

    ausgabe = args.ausgabe or args.csv.with_name(args.csv.stem + "_ergaenzt.csv")
    df.to_csv(ausgabe, sep=";", encoding=ENCODING, index=False, lineterminator="\r\n",
              quoting=csv.QUOTE_MINIMAL)
    print(f"Artikelliste geschrieben: {ausgabe} ({len(df)} Artikel, {len(df.columns)} Spalten)")


if __name__ == "__main__":
    main()
