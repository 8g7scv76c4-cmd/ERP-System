"""Ergänzt die OrgaCalc-Artikelliste (CSV aus dem Zeichenprogramm) um die
Spalten Länge, Breite, Dicke und Verschnitt.

Aufruf:
    python artikel_ergaenzen.py <Artikel.csv> [--regeln Verschnitt_Regeln.xlsx]

Ohne --regeln wird zusätzlich eine Vorlage "Verschnitt_Regeln.xlsx" erzeugt:
eine Zeile pro Hauptgruppe/Gruppe, in die man den Verschnitt einträgt.
Optional kann in der Spalte "Stichwort" ein Text stehen; die Regel gilt dann
nur für Artikel, deren RWDM-Nr., Beschreibung oder Katalogpfad ihn enthält
(z. B. Gruppe "Blum Legrabox" + Stichwort "Zargen"). Regeln mit Stichwort
haben Vorrang vor Regeln ohne; passen mehrere, gewinnt die obere.

Länge, Breite und Dicke (in mm) werden aus der Beschreibung gelesen
("Länge 1200mm", "NL=450", "30x3mm", "... 2500mm"). Kleinteile wie Topfbänder,
Schlösser und Schrauben (siehe OHNE_MASSE) bleiben leer. Jeder Fund steht mit
Fundstelle in "Masse_Kontrolle.xlsx" zum Nachprüfen.

Die Ausgabe-CSV hat dasselbe Format wie die Eingabe (Semikolon, Windows-1252,
CRLF), damit OrgaCalc sie wie gewohnt einlesen kann. Alle Werte bleiben Text,
Bestellnummern werden also nicht verändert.
"""

import argparse
import csv
import re
from pathlib import Path

import pandas as pd

ENCODING = "cp1252"
NEUE_SPALTEN = ["Länge", "Breite", "Dicke", "Verschnitt"]
REGEL_SPALTEN = ["Hauptgruppe Name", "Gruppe Name", "Stichwort", "Verschnitt"]

# Kleinteile, bei denen OrgaCalc keine Masse braucht (Topfbänder, Schlösser, Schrauben ...)
OHNE_MASSE = {
    "Moebelband", "Tuerband", "Tuerbeschlag", "Moebelschliessung", "Verbindungen",
    "M-Schrauben", "Montage-Schrauben", "Puffer", "Magnet", "Winkelverbinder",
    "Verstaerkungswinkel", "Ausrichtbeschlag", "Rollen", "Moebeloeffnung", "Antrieb",
    "Klappe", "Armaturen", "Becken",
}
# Profile/Rohre: "AxB" ist der Querschnitt, die Länge kommt aus dem Projekt
QUERSCHNITT_GRUPPEN = {"MetallProfil", "AluProfil", "Schlitzrohr T30", "Schlitzrohr T35", "Schlitzrohr T50"}

ZAHL = r"(\d+(?:[.,]\d+)?)"
EINHEIT = r"\s*(mm|cm|m)\b"
RE_LAENGE = re.compile(
    r"(?:\b(?:Lagerlänge|Länge|NL|L)\s*[:=]?\s*)" + ZAHL + r"(?:" + EINHEIT + r")?(?![\d-])",
    re.IGNORECASE,
)
RE_BREITE = re.compile(r"\b(?:B|Breite)\s*[:=]?\s*" + ZAHL + EINHEIT + r"|" + ZAHL + EINHEIT + r"\s*Breit",
                       re.IGNORECASE)
RE_DICKE = re.compile(r"\b(?:Dicke|Stärke)\s*[:=]?\s*" + ZAHL + EINHEIT + r"|" + ZAHL + EINHEIT + r"\s*Dick",
                      re.IGNORECASE)
RE_QUER = re.compile(r"(?<![M\d.,])" + ZAHL + r"\s*(?:mm)?\s*x\s*" + ZAHL + r"(?:\s*x\s*" + ZAHL + r")?" + EINHEIT, re.IGNORECASE)
# Zahl mit Einheit ohne Kennwort davor, z. B. "Griffleiste ... silber 2500mm"
RE_NACKT = re.compile(r"(?<![\d.,x/-])" + ZAHL + EINHEIT + r"(?![\w-])", re.IGNORECASE)
# Kennwörter, deren Zahl kein Artikelmass ist (Dornmass, Lochabstand, Korpusbreite ...)
KEIN_MASS = re.compile(
    r"^(?:DM|Dis|N|Stulp|Eckstulp|LA|Lochabstand|KB|Korpusbreite|Plattendicke|Materialstärke|Lochteil|"
    r"Ausladung|für|Ø|ø|D|H|Höhe|Profilhöhe|Tief|Türdicke|Türendicke|Lagerlänge|Länge|NL|L|B|Breite|"
    r"Dicke|Stärke|Türstärke|Se|WS|T|M\d*|[\d.,]*x)[:=]?$",
    re.IGNORECASE,
)


def _mm(zahl, einheit):
    wert = float(zahl.replace(",", "."))
    wert *= {"m": 1000, "cm": 10}.get((einheit or "mm").lower(), 1)
    return f"{wert:g}"


def masse_aus_beschreibung(text, hauptgruppe="", gruppe=""):
    """Liest Länge, Breite und Dicke (in mm) aus der Beschreibung.

    Gibt (länge, breite, dicke, fundstelle) zurück; nicht gefundene Werte sind "".
    """
    if hauptgruppe in OHNE_MASSE:
        return "", "", "", ""
    laenge = breite = dicke = ""
    funde = []

    for m in RE_LAENGE.finditer(text):
        # ohne Einheit nur ab 3 Stellen ("NL450", "L0300"), sonst z. B. Grösse "L7"
        if m.group(2) or len(m.group(1)) >= 3:
            laenge = _mm(m.group(1), m.group(2))
            funde.append(m.group(0).strip())
            break
    if m := RE_BREITE.search(text):
        breite = _mm(m.group(1) or m.group(3), m.group(2) or m.group(4))
        funde.append(m.group(0).strip())
    if m := RE_DICKE.search(text):
        dicke = _mm(m.group(1) or m.group(3), m.group(2) or m.group(4))
        funde.append(m.group(0).strip())

    rest = text
    if m := RE_QUER.search(text):
        rest = text[: m.start()] + " " + text[m.end():]
        a, b, c, einheit = m.group(1), m.group(2), m.group(3), m.group(4)
        werte = [_mm(w, einheit) for w in (a, b, c) if w]
        if gruppe in QUERSCHNITT_GRUPPEN or hauptgruppe in QUERSCHNITT_GRUPPEN:
            # Querschnitt: Breite x Höhe (x Wandstärke) -> Breite, Dicke
            breite, dicke = breite or werte[0], dicke or werte[1]
        elif len(werte) == 3:
            laenge, breite, dicke = laenge or werte[0], breite or werte[1], dicke or werte[2]
        elif float(werte[1]) <= 20:
            breite, dicke = breite or werte[0], dicke or werte[1]
        else:
            breite, laenge = breite or werte[0], laenge or werte[1]
        funde.append(m.group(0).strip())

    if not laenge:
        # letzte "nackte" Zahl mit Einheit, z. B. Profillänge am Textende
        for m in reversed(list(RE_NACKT.finditer(rest))):
            davor = rest[: m.start()].split()
            wort = davor[-1].rstrip(",") if davor else ""
            wert = _mm(m.group(1), m.group(2))
            # kleine Zahlen sind meist Stärken/Durchmesser, keine Länge
            if not KEIN_MASS.match(wort) and float(wert) >= 50:
                laenge = wert
                funde.append(m.group(0).strip())
                break

    return laenge, breite, dicke, " | ".join(funde)


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
    # Regeln mit Stichwort zuerst, damit sie Vorrang haben; sonst gilt die
    # Reihenfolge in der Tabelle (erste passende Regel gewinnt)
    return regeln.sort_values("Stichwort", key=lambda s: s == "", kind="stable")


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

    masse = df.apply(
        lambda r: masse_aus_beschreibung(r["Beschreibung"], r["Hauptgruppe Name"], r["Gruppe Name"]),
        axis=1, result_type="expand",
    )
    masse.columns = ["Länge", "Breite", "Dicke", "Fundstelle"]
    for spalte in ("Länge", "Breite", "Dicke"):
        leer = df[spalte] == ""
        df.loc[leer, spalte] = masse.loc[leer, spalte]
    gefunden = masse["Fundstelle"] != ""
    kontrolle = pd.concat(
        [df.loc[gefunden, ["Id", "RWDM-Nr.", "Hauptgruppe Name", "Gruppe Name", "Beschreibung"]],
         masse.loc[gefunden]], axis=1,
    )
    kontrolle_pfad = args.csv.with_name("Masse_Kontrolle.xlsx")
    with pd.ExcelWriter(kontrolle_pfad) as writer:
        kontrolle.to_excel(writer, sheet_name="Masse", index=False)
        ws = writer.book["Masse"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for spalte, breite in zip("ABCDEFGHI", (7, 30, 18, 24, 80, 9, 9, 9, 40)):
            ws.column_dimensions[spalte].width = breite
    print(f"Masse gefunden bei {gefunden.sum()} Artikeln, Kontrollliste: {kontrolle_pfad}")

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
