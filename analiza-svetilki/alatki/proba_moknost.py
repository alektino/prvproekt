"""Проба: колкава моќност треба да имаат старите светилки?

Обратно од analiza.py: бројот на стари светилки е фиксиран на вашата проценка
(1.500), а се пресметува колкава моќност треба да има секоја стара светилка за
да ја објасни енергијата што не ја трошат LED светилките од ЈПП.

  моќност (W) = (kWh × (1 − загуби) − kWh на LED) × 1000 ÷ (број стари × удел активни × ноќни часови)

Две основи за kWh:
  A – сите фактурирани kWh без новогодишното украсување
  B – само ноќната енергија (без постојаниот 24-часовен товар и без украсувањето)
Загубите се загубите по броилото (кабли до светилките, уреди во ормарите).

Употреба (по analiza.py):  python3 analiza-svetilki/alatki/proba_moknost.py
Излез: izlez/proba_moknost.csv, izlez/proba_moknost_mesecno.csv и лист „Проба – моќност“
во izlez/Анализа_светилки.xlsx
"""
import csv
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

import analiza as A

LOSSES = [0.0, 0.05, 0.10]
ACTIVE = [0.8, 0.9, 1.0]
BASES = {
    "A": ("Сите kWh без украсување", lambda r: r["kwh"] - r["kwh_ukras"]),
    "B": ("Само ноќна енергија (без постојан товар и украсување)", lambda r: r["kwh_nokno"]),
}
MAIN_LOSS = 0.10


def main():
    a, months, rows, by_group, compare, profiles = A.analyse()
    total_old = float(a["stari_procenka_vkupno"])
    hours = {m: next(r["nokni_casovi"] for r in rows if r["mesec"] == m) for m in months}
    hours["all"] = sum(hours.values())
    guess = {c["opstina"]: c["procenka_stari"] for c in compare if c["opstina"] != "Вкупно"}
    scale = total_old / sum(guess.values())
    old_count = {g: n * scale for g, n in guess.items()}
    old_count["Вкупно"] = total_old
    g_row = {(r["opstina"], r["mesec"]): r for r in by_group}

    def watt(r, base, loss, active, n, h):
        outside = BASES[base][1](r) * (1 - loss) - r["kwh_led"]
        return outside, outside * 1000 / (n * active * h) if n else None

    grid = []
    for g in A.GROUPS + ["Вкупно"]:
        r = g_row[g, "all"]
        for base in BASES:
            for loss in LOSSES:
                outside, _ = watt(r, base, loss, 1, old_count[g], hours["all"])
                row = {"opstina": g, "osnova": base, "osnova_opis": BASES[base][0], "zagubi": loss, "led": r["led"],
                       "stari": round(old_count[g]), "kwh_osnova": round(BASES[base][1](r)), "kwh_led": round(r["kwh_led"]),
                       "kwh_nadvor_od_led": round(outside)}
                for act in ACTIVE:
                    row[f"w_aktivni_{int(act * 100)}"] = round(watt(r, base, loss, act, old_count[g], hours["all"])[1], 1)
                grid.append(row)

    monthly = []
    for m in months:
        r = g_row["Вкупно", m]
        row = {"mesec": m, "nokni_casovi": round(hours[m], 1), "kwh": round(r["kwh"]), "kwh_ukras": round(r["kwh_ukras"]),
               "kwh_postojano": round(r["kwh_postojano"]), "kwh_nokno": round(r["kwh_nokno"]), "kwh_led": round(r["kwh_led"])}
        for base in BASES:
            for act in (0.8, 1.0):
                row[f"w_{base}_{int(act * 100)}"] = round(watt(r, base, MAIN_LOSS, act, total_old, hours[m])[1], 1)
        monthly.append(row)

    for name, data in (("proba_moknost.csv", grid), ("proba_moknost_mesecno.csv", monthly)):
        with open(A.IZLEZ / name, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=list(data[0].keys()))
            w.writeheader(); w.writerows(data)

    write_sheet(a, grid, monthly, total_old)
    for r in grid:
        if r["opstina"] == "Вкупно":
            print(f"{r['osnova']} загуби {r['zagubi']:.0%}: надвор од LED {r['kwh_nadvor_od_led']:8,} kWh → W при 80/90/100 %: "
                  f"{r['w_aktivni_80']} / {r['w_aktivni_90']} / {r['w_aktivni_100']}")


def write_sheet(a, grid, monthly, total_old):
    path = A.IZLEZ / "Анализа_светилки.xlsx"
    wb = load_workbook(path)
    if "Проба – моќност" in wb.sheetnames:
        del wb["Проба – моќност"]
    ws = wb.create_sheet("Проба – моќност", 1)
    bold, head = Font(bold=True), PatternFill("solid", fgColor="E8EEF2")
    hl = PatternFill("solid", fgColor="E3F1E3")
    ws.append(["Проба: колкава моќност треба да имаат " + f"{total_old:,.0f}".replace(",", ".") + " стари светилки за да ја објасни енергијата надвор од LED?"])
    ws["A1"].font = Font(bold=True, size=13)
    ws.append(["Моќност (W) = (kWh × (1 − загуби) − kWh на LED) × 1000 ÷ (број стари × удел активни × ноќни часови). LED светат цела ноќ со номиналната моќност."])
    ws.append(["A = сите фактурирани kWh без новогодишното украсување; B = само ноќната енергија (без постојаниот 24-часовен товар и украсувањето)."])
    ws.append(["Загуби = загуби по броилото (кабли од ормарот до светилките, уреди во ормарите). Загубите пред броилото не се фактурираат."])
    ws.append(["По поранешна општина: 1.500 стари светилки распоредени според вашите проценки за удел."])
    ws.append([])
    cols = ["Поранешна општина", "Основа", "Загуби", "LED", "Стари", "kWh основа", "kWh LED", "kWh надвор од LED", "W при 80 % активни", "W при 90 % активни", "W при 100 % активни"]
    ws.append(cols)
    for c in ws[ws.max_row]:
        c.font, c.fill, c.alignment = bold, head, Alignment(wrap_text=True, vertical="top")
    for r in grid:
        ws.append([r["opstina"], r["osnova"], r["zagubi"], r["led"], r["stari"], r["kwh_osnova"], r["kwh_led"], r["kwh_nadvor_od_led"],
                   r["w_aktivni_80"], r["w_aktivni_90"], r["w_aktivni_100"]])
        row = ws[ws.max_row]
        row[2].number_format = "0%"
        for c in row[3:8]:
            c.number_format = "#,##0"
        if r["opstina"] == "Вкупно":
            for c in row:
                c.font = bold
            if r["osnova"] == "B" and r["zagubi"] == MAIN_LOSS:
                for c in row:
                    c.fill = hl
    ws.append([])
    ws.append(["По месец (сите општини, загуби " + f"{MAIN_LOSS:.0%}" + ", " + f"{total_old:,.0f}".replace(",", ".") + " стари светилки)"])
    ws[ws.max_row][0].font = bold
    mcols = ["Месец", "Ноќни часови", "kWh", "Украсување", "Постојан товар", "Ноќна енергија", "kWh LED", "W (A, 80 %)", "W (A, 100 %)", "W (B, 80 %)", "W (B, 100 %)"]
    ws.append(mcols)
    for c in ws[ws.max_row]:
        c.font, c.fill, c.alignment = bold, head, Alignment(wrap_text=True, vertical="top")
    for r in monthly:
        ws.append([r["mesec"], r["nokni_casovi"], r["kwh"], r["kwh_ukras"], r["kwh_postojano"], r["kwh_nokno"], r["kwh_led"], r["w_A_80"], r["w_A_100"], r["w_B_80"], r["w_B_100"]])
        for c in ws[ws.max_row][2:7]:
            c.number_format = "#,##0"
    for col, w in zip("ABCDEFGHIJK", [18, 8, 9, 8, 8, 11, 11, 13, 11, 11, 11]):
        ws.column_dimensions[col].width = w
    wb.save(path)


if __name__ == "__main__":
    main()
