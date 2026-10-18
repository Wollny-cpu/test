"""Erzeugt die DIN-277-Flächenauswertung aus den Daten des Repositories cafm-acw.

Quellen:
  A) backend/seed/app_state_live.json  (CAFM-Datenstand, Mandant TEN-001)
  B) docs/test-results/F3.2_AEC_Muster_Ergebnis.json  (AEC-Import EG, DE-SIE-11154)

Aufruf: python3 din277_auswertung.py <pfad-zu-cafm-acw> <ausgabe.xlsx>
"""
import json
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

repo, out = Path(sys.argv[1]), Path(sys.argv[2])

# ---------------------------------------------------------------- Zuordnung
# Raumbezeichnung / Nutzungsart -> (DIN-277:2021-Code, Grundlage, prüfen?)
MAP = {
    "büro": ("NUF 2.1", False),
    "besprechung": ("NUF 2.3", False),
    "büro (registratur)": ("NUF 2.1", False),
    "büro(poststelle)": ("NUF 2.1", True),
    "büro(plank/archiv/kop.)": ("NUF 2.1", True),
    "büro (lager)": ("NUF 4.1", True),
    "aufenthaltsraum": ("NUF 1.3", False),
    "cafeteria / casino": ("NUF 1.5", False),
    "gäste 1": ("NUF 1.5", True),
    "gäste 2": ("NUF 1.5", True),
    "kochstudio": ("NUF 3.7", True),
    "koch- und bratküche": ("NUF 3.7", False),
    "vorbereitungsküche": ("NUF 3.7", False),
    "geschirrspüle": ("NUF 3.7", False),
    "gästeanrichte": ("NUF 3.7", False),
    "funkwerkstatt": ("NUF 3.2", False),
    "lichtstudio": ("NUF 4.6", True),
    "tageslager": ("NUF 4.1", False),
    "anlieferung": ("NUF 4.4", False),
    "an- u. rücklieferung": ("NUF 4.4", False),
    "kundenforum": ("NUF 5.6", True),
    "werksarzt": ("NUF 6.1", False),
    "werksarzt sprechzimmer": ("NUF 6.1", False),
    "dusche": ("NUF 7.1", False),
    "waschen": ("NUF 7.1", False),
    "stiefel waschen": ("NUF 7.1", False),
    "beh. wc": ("NUF 7.1", False),
    "garderobe": ("NUF 7.2", False),
    "technik elv": ("TF 8", False),
    "fm-technik": ("TF 8", True),
    "tel": ("TF 8", True),
    "empfang / foyer": ("VF 9.1", True),
}


def fix(s):  # CP850-Fehlkodierung aus der DWG (³ = ü, õ = ä)
    return (s or "").replace("³", "ü").replace("õ", "ä").strip()


def classify(label):
    key = fix(label).lower()
    if not key:
        return "offen", "keine Bezeichnung", True
    if key in MAP:
        code, check = MAP[key]
        return code, "Bezeichnung", check
    return "offen", "Bezeichnung nicht eindeutig", True


rows = []  # Quelle, Gebäude, Geschoss, Raumnr, Bezeichnung, Nutzungsart, NRF, Code, (Hauptgruppe), Grundlage, prüfen, Bemerkung

# A) CAFM-Datenstand
state = json.loads((repo / "backend/seed/app_state_live.json").read_text("utf-8"))["state"]
for tid, t in state["tenantData"].items():
    objs = {o["id"]: o for o in t["objects"]}
    types = {ty["id"]: ty for ty in t["types"]}
    for o in t["objects"]:
        ty = types.get(o["typeId"])
        if not ty or ty["name"].strip() != "Raum":
            continue
        attr = {a["name"]: a["id"] for a in ty["attrs"]}
        floor = objs.get(o["parentId"])
        bld = objs.get(floor["parentId"]) if floor else None
        fl = f'{floor["name"]} / {floor["number"]}' if floor else "?"
        nutz = o["values"].get(attr.get("Nutzungsart"), "")
        raw = o["values"].get(attr.get("Netto-Raumfläche"), "")
        area = float(raw) if raw else None
        code, basis, check = classify(nutz or o["name"])
        if nutz:
            basis = "CAFM-Nutzungsart"
        note = "" if area is not None else "Netto-Raumfläche fehlt"
        rows.append([f"CAFM {tid}", f'{bld["name"] if bld else "?"} ({tid})', fl, o["number"], o["name"],
                     nutz, area, code, None, basis, "ja" if (check or area is None) else "nein", note])

# B) AEC-Import EG
aec = json.loads((repo / "docs/test-results/F3.2_AEC_Muster_Ergebnis.json").read_text("utf-8"))
bld_aec = "DE-SIE-11154 (EG-Zeichnung)"
for s in aec["spaces"]:
    props = {p["name"]: p["displayValue"] for ps in s["propertySets"] for p in ps["properties"]
             if p["displayValue"]}
    code, basis, check = classify(s["name"])
    note = "Raumnummer fehlt in DWG" if props.get("Nummer") == "??" else ""
    rows.append(["DWG-Import F3.2", bld_aec, "EG", props.get("Nummer", ""), fix(s["name"]), "",
                 round(s["aecArea"], 2), code, None, basis, "ja" if check else "nein", note])

# ---------------------------------------------------------------- Workbook
F = "Arial"
hdr_font = Font(name=F, bold=True, color="FFFFFF")
hdr_fill = PatternFill("solid", fgColor="305496")
sub_fill = PatternFill("solid", fgColor="D9E1F2")
tot_fill = PatternFill("solid", fgColor="BDD7EE")
inp_fill = PatternFill("solid", fgColor="FFF2CC")
thin = Side(style="thin", color="A6A6A6")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
base = Font(name=F)
blue = Font(name=F, color="0000FF")
bold = Font(name=F, bold=True)
M2 = '#,##0.00;-#,##0.00;"-"'

wb = Workbook()
ws_a = wb.active
ws_a.title = "Auswertung"
ws_r = wb.create_sheet("Raumliste")
ws_k = wb.create_sheet("DIN277-Katalog")
ws_h = wb.create_sheet("Hinweise")

# Raumliste
head = ["Quelle", "Gebäude", "Geschoss", "Raumnr.", "Bezeichnung", "Nutzungsart (CAFM)",
        "NRF [m²]", "DIN 277 Code", "Hauptgruppe", "Zuordnung über", "prüfen", "Bemerkung"]
ws_r.append(head)
for r in rows:
    ws_r.append(r)
n = len(rows) + 1
for i in range(2, n + 1):
    ws_r[f"I{i}"] = (f'=IF(LEFT(H{i},3)="NUF",LEFT(H{i},5),IF(LEFT(H{i},2)="TF","TF",'
                     f'IF(LEFT(H{i},2)="VF","VF","offen")))')
for row in ws_r.iter_rows(min_row=1, max_row=n):
    for c in row:
        c.font, c.border = base, box
for c in ws_r[1]:
    c.font, c.fill = hdr_font, hdr_fill
    c.alignment = Alignment(wrap_text=True, vertical="center")
for i in range(2, n + 1):
    ws_r[f"G{i}"].number_format = M2
    ws_r[f"H{i}"].font = blue
    ws_r[f"H{i}"].fill = inp_fill
    if ws_r[f"K{i}"].value == "ja":
        ws_r[f"K{i}"].font = Font(name=F, bold=True, color="C00000")
for col, w in zip("ABCDEFGHIJKL", [16, 44, 26, 11, 26, 18, 11, 13, 12, 24, 8, 26]):
    ws_r.column_dimensions[col].width = w
ws_r.freeze_panes = "A2"
ws_r.auto_filter.ref = f"A1:L{n}"
ws_r["H1"].comment = Comment("Eingabespalte (gelb): DIN-277-Code anpassen, z. B. 'NUF 2.1', "
                             "'TF 8', 'VF 9.1' oder 'offen'. Die Auswertung rechnet automatisch neu.",
                             "Auswertung")

# Katalog (Liste für Dropdown)
cat = [("NUF 1.1", "Wohnräume"), ("NUF 1.2", "Gemeinschaftsräume"), ("NUF 1.3", "Pausenräume"),
       ("NUF 1.4", "Warteräume"), ("NUF 1.5", "Speiseräume"),
       ("NUF 2.1", "Büroräume"), ("NUF 2.2", "Großraumbüros"), ("NUF 2.3", "Besprechungsräume"),
       ("NUF 3.2", "Werkstätten"), ("NUF 3.7", "Küchen"), ("NUF 3.8", "Sonderarbeitsräume"),
       ("NUF 4.1", "Lagerräume"), ("NUF 4.2", "Archive, Sammlungsräume"),
       ("NUF 4.4", "Annahme- und Ausgaberäume"), ("NUF 4.6", "Ausstellungsräume"),
       ("NUF 5.6", "Versammlungsräume"), ("NUF 6.1", "Räume mit allg. medizinischer Ausstattung"),
       ("NUF 7.1", "Sanitärräume"), ("NUF 7.2", "Garderoben"), ("NUF 7.3", "Abstellräume"),
       ("TF 8", "Technikfläche"), ("VF 9.1", "Flure, Hallen"), ("VF 9.2", "Treppen"),
       ("VF 9.3", "Schächte für Förderanlagen"), ("offen", "noch nicht zugeordnet")]
ws_k.append(["Code", "Bezeichnung (DIN 277:2021, Auszug)"])
for c in cat:
    ws_k.append(list(c))
for row in ws_k.iter_rows():
    for c in row:
        c.font, c.border = base, box
for c in ws_k[1]:
    c.font, c.fill = hdr_font, hdr_fill
ws_k.column_dimensions["A"].width, ws_k.column_dimensions["B"].width = 10, 46
dv = DataValidation(type="list", formula1=f"='DIN277-Katalog'!$A$2:$A${len(cat) + 1}",
                    allow_blank=False, showErrorMessage=False)
ws_r.add_data_validation(dv)
dv.add(f"H2:H{n}")

# Auswertung
groups = ["NUF 1", "NUF 2", "NUF 3", "NUF 4", "NUF 5", "NUF 6", "NUF 7"]
ws_a["A1"] = "Flächenauswertung nach DIN 277:2021 – alle Geschosse"
ws_a["A1"].font = Font(name=F, bold=True, size=14)
ws_a["A2"] = ("Netto-Raumflächen in m², Bereich a (regulär umschlossen). Quelle: Repository "
              "cafm-acw, siehe Blatt „Hinweise“.")
ws_a["A2"].font = Font(name=F, italic=True, color="595959")
cols = (["Gebäude", "Geschoss", "Räume"] + groups +
        ["NUF gesamt", "TF", "VF", "NRF zugeordnet", "nicht zugeordnet", "NRF gesamt",
         "Zuordnungs-quote", "BGF", "KGF"])
HR = 4
for j, h in enumerate(cols, 1):
    c = ws_a.cell(HR, j, h)
    c.font, c.fill, c.border = hdr_font, hdr_fill, box
    c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
ws_a.row_dimensions[HR].height = 32

R = "Raumliste!"
G, C, A, I = f"{R}$B$2:$B${n}", f"{R}$C$2:$C${n}", f"{R}$G$2:$G${n}", f"{R}$I$2:$I${n}"
blds = []
for r in rows:
    if r[1] not in blds:
        blds.append(r[1])
r_i = HR + 1
total_rows = []
for b in blds:
    floors = []
    for r in rows:
        if r[1] == b and r[2] not in floors:
            floors.append(r[2])
    first = r_i
    for fl in floors:
        ws_a.cell(r_i, 1, b)
        ws_a.cell(r_i, 2, fl)
        crit = f'{G},$A{r_i},{C},$B{r_i}'
        ws_a.cell(r_i, 3, f"=COUNTIFS({crit})")
        for j, g in enumerate(groups + [None, "TF", "VF"], 4):
            if g is None:
                ws_a.cell(r_i, j, f"=SUM(D{r_i}:J{r_i})")
            else:
                ws_a.cell(r_i, j, f'=SUMIFS({A},{crit},{I},"{g}")')
        ws_a.cell(r_i, 14, f"=K{r_i}+L{r_i}+M{r_i}")
        ws_a.cell(r_i, 15, f'=SUMIFS({A},{crit},{I},"offen")')
        ws_a.cell(r_i, 16, f"=N{r_i}+O{r_i}")
        ws_a.cell(r_i, 17, f"=IF(P{r_i}=0,0,N{r_i}/P{r_i})")
        ws_a.cell(r_i, 18, "n. e.")
        ws_a.cell(r_i, 19, "n. e.")
        r_i += 1
    ws_a.cell(r_i, 1, b)
    ws_a.cell(r_i, 2, "Summe Gebäude")
    for j in range(3, 17):
        L = get_column_letter(j)
        ws_a.cell(r_i, j, f"=SUM({L}{first}:{L}{r_i - 1})")
    ws_a.cell(r_i, 17, f"=IF(P{r_i}=0,0,N{r_i}/P{r_i})")
    ws_a.cell(r_i, 18, "n. e.")
    ws_a.cell(r_i, 19, "n. e.")
    for j in range(1, 20):
        ws_a.cell(r_i, j).fill = sub_fill
        ws_a.cell(r_i, j).font = bold
    total_rows.append(r_i)
    r_i += 1
ws_a.cell(r_i, 1, "Gesamt alle Gebäude / Geschosse")
for j in range(3, 17):
    L = get_column_letter(j)
    ws_a.cell(r_i, j, "=" + "+".join(f"{L}{t}" for t in total_rows))
ws_a.cell(r_i, 17, f"=IF(P{r_i}=0,0,N{r_i}/P{r_i})")
ws_a.cell(r_i, 18, "n. e.")
ws_a.cell(r_i, 19, "n. e.")
grand = r_i
for row in ws_a.iter_rows(min_row=HR + 1, max_row=grand, max_col=19):
    for c in row:
        if c.font != bold:
            c.font = base
        c.border = box
        if 4 <= c.column <= 16:
            c.number_format = M2
        if c.column == 17:
            c.number_format = '0.0%'
        if c.column in (18, 19):
            c.alignment = Alignment(horizontal="center")
            c.font = Font(name=F, color="808080")
for j in range(1, 20):
    ws_a.cell(grand, j).fill = tot_fill
    ws_a.cell(grand, j).font = bold
ws_a.cell(HR, 18).comment = Comment("n. e. = nicht ermittelbar: Die Datenquellen enthalten nur "
                                    "Raumpolygone, keine Geschossumrisse oder Konstruktionsflächen.",
                                    "Auswertung")

# Kennzahlen je Gebäude
k = grand + 3
ws_a.cell(k - 1, 1, "Kennzahlen je Gebäude").font = Font(name=F, bold=True, size=12)
kh = ["Gebäude", "NUF / NRF", "VF / NUF", "TF / NRF", "Anteil NUF 2 Büro an NUF",
      "Räume zu prüfen"]
for j, h in enumerate(kh, 1):
    c = ws_a.cell(k, j, h)
    c.font, c.fill, c.border = hdr_font, hdr_fill, box
    c.alignment = Alignment(wrap_text=True, horizontal="center")
for t, b in zip(total_rows, blds):
    k += 1
    ws_a.cell(k, 1, f"=A{t}")
    ws_a.cell(k, 2, f"=IF(N{t}=0,0,K{t}/N{t})")
    ws_a.cell(k, 3, f"=IF(K{t}=0,0,M{t}/K{t})")
    ws_a.cell(k, 4, f"=IF(N{t}=0,0,L{t}/N{t})")
    ws_a.cell(k, 5, f"=IF(K{t}=0,0,E{t}/K{t})")
    ws_a.cell(k, 6, f'=COUNTIFS({G},A{t},{R}$K$2:$K${n},"ja")')
    for j in range(1, 7):
        c = ws_a.cell(k, j)
        c.font, c.border = base, box
        if 2 <= j <= 5:
            c.number_format = "0.0%"
ws_a.cell(k + 1, 1, "Bezugsgröße der Quoten: zugeordnete NRF (nicht zugeordnete Räume bleiben "
                    "außen vor).").font = Font(name=F, italic=True, color="595959")

widths = [44, 24, 7] + [10] * 13 + [10, 7, 7]
for j, w in enumerate(widths, 1):
    ws_a.column_dimensions[get_column_letter(j)].width = w
ws_a.freeze_panes = ws_a.cell(HR + 1, 3)

# Hinweise
notes = [
    ("Gegenstand", "Flächenauswertung nach DIN 277:2021 (Grundflächen und Rauminhalte im Hochbau) "
     "über alle Geschosse, die im Repository cafm-acw mit Flächendaten vorliegen."),
    ("Quelle A", "backend/seed/app_state_live.json – CAFM-Datenstand, Mandant TEN-001 "
     "(Testliegenschaft Alpha / Testgebäude A, Geschosse 01 und 02). Fläche = Merkmal "
     "„Netto-Raumfläche“, Zuordnung über Merkmal „Nutzungsart“. TEN-002 enthält nur einen "
     "Testraum ohne Fläche, TEN-003 keine Objekte."),
    ("Quelle B", "docs/test-results/F3.2_AEC_Muster_Ergebnis.json – Ergebnis des AEC-Flächenimports "
     "der Zeichnung DE-SIE-11154-04-02-01_EG (124 AecDbSpaces, Layer A_Raeume). Fläche = aecArea. "
     "In der Zeichnung liegt nur das Erdgeschoss vor."),
    ("Zuordnung", "DIN-277-Codes wurden aus Raumbezeichnung bzw. Nutzungsart abgeleitet. "
     "Untergruppen (z. B. NUF 2.1) sind Vorschläge; Hauptgruppen (NUF 1–7, TF, VF) sind für die "
     "Auswertung maßgeblich. Räume mit unklarer Bezeichnung sind in der Raumliste mit "
     "„prüfen = ja“ markiert."),
    ("Nicht zugeordnet", "Der Großteil der DWG-Räume hat keine Raumbezeichnung (Property „Name“ leer). "
     "Diese Räume stehen auf „offen“ und erscheinen in der Spalte „nicht zugeordnet“. Sobald der "
     "Code in Raumliste!H gesetzt wird, rechnet die Auswertung neu."),
    ("BGF / KGF", "Nicht ermittelbar (n. e.): Beide Quellen enthalten nur Raumflächen (NRF), keine "
     "Geschossumrisse, Wand- oder Konstruktionsflächen. Für BGF = NRF + KGF wird ein BGF-Polygon je "
     "Geschoss (DWG-Layer oder CAFM-Merkmal) benötigt."),
    ("Bereiche a/b/c", "Alle Flächen wurden als Bereich a (regulär umschlossen) gewertet; "
     "Balkone, Loggien o. Ä. sind in den Quellen nicht gekennzeichnet."),
    ("Datenqualität", "Geschoss OBJ-031: Name „02“ und Nummer „1.Obergeschoss“ sind im CAFM "
     "vertauscht gepflegt. Raum 101 (Testraum) hat keine Netto-Raumfläche. DWG-Raum 0069 hat die "
     "Nummer „??“. Das DWG-Property „OKFFB“ enthält Ziffern 1–9, die keiner DIN-277-Gruppe "
     "konsistent entsprechen (z. B. Büros = 1) und wurden daher nicht verwendet. Umlaute aus der "
     "DWG (CP850) wurden korrigiert."),
    ("Systemabgleich", "Das CAFM ACW kennt DIN 277 bisher als zentral gepflegten Systemkatalog, als "
     "Merkmal „DIN-277-Kategorie“ (Freitext beim BIM-Import) und für Piktogramme. Eine fertige "
     "DIN-277-Auswertung je Geschoss gibt es im System noch nicht; das Mandantenmerkmal "
     "„Nutzungsart“ nutzt aktuell einen Testkatalog (Büro, Besprechung, Technik)."),
    ("Bedienung", "Gelbe Zellen in Raumliste!H sind die einzigen Eingabezellen (Dropdown aus "
     "Blatt „DIN277-Katalog“). Alle Summen und Kennzahlen sind Formeln."),
]
ws_h.append(["Thema", "Hinweis"])
for t in notes:
    ws_h.append(list(t))
for row in ws_h.iter_rows():
    for c in row:
        c.font, c.border = base, box
        c.alignment = Alignment(wrap_text=True, vertical="top")
for c in ws_h[1]:
    c.font, c.fill = hdr_font, hdr_fill
ws_h.column_dimensions["A"].width, ws_h.column_dimensions["B"].width = 18, 110

wb.calculation.fullCalcOnLoad = True
wb.save(out)
print(f"{len(rows)} Räume, {len(blds)} Gebäude -> {out}")
