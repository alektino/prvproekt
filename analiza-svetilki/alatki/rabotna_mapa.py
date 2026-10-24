"""Работна табела за проверка на поврзувањето мерно место → населено место.

  python3 analiza-svetilki/alatki/rabotna_mapa.py izvezi
      Ја прави izlez/Проверка_мерни_места.xlsx од vlez/mapiranje.csv.

  python3 analiza-svetilki/alatki/rabotna_mapa.py uvezi <поправена_табела.xlsx>
      Ги зема поправките (колони „Точно?“ и „Точно населено место“) и го
      ажурира vlez/mapiranje.csv. Потоа повторно пуштете analiza.py.
"""
import csv
import sys
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parent.parent
VLEZ, IZLEZ = ROOT / "vlez", ROOT / "izlez"
GROUPS = ["Кичево", "Другово", "Осломеј", "Вранештица", "Зајас"]
METER_PREFIX = {"21021": "Кичево", "210316": "Другово", "210365": "Вранештица", "210366": "Зајас", "210367": "Осломеј"}
NOT_LIGHTING = "Не е улично осветлување"
HEAD, FILL = Font(bold=True), PatternFill("solid", fgColor="E8EEF2")
TODO = PatternFill("solid", fgColor="FFF4CC")
CHECK = PatternFill("solid", fgColor="FDE2E1")
FIELDS = ["meter", "address", "settlement", "unit", "group", "certainty"]


def read(name):
    return list(csv.DictReader(open(VLEZ / name, encoding="utf-8")))


def group_of_meter(meter):
    return next((g for p, g in METER_PREFIX.items() if meter.startswith(p)), "")


def header(ws, cols, widths):
    ws.append(cols)
    for c in ws[1]:
        c.font, c.fill, c.alignment = HEAD, FILL, Alignment(wrap_text=True, vertical="top")
    for i, w in enumerate(widths):
        ws.column_dimensions[chr(65 + i)].width = w
    ws.freeze_panes = "A2"


def izvezi():
    places, lamps, mapping = read("naseleni.csv"), read("svetilki.csv"), read("mapiranje.csv")
    kwh, addresses = Counter(), defaultdict(set)
    for r in read("stavki.csv"):
        if "LV1" in r["category"]:
            kwh[r["meter"]] += float(r["kwh"])
            addresses[r["meter"]].add(r["address"].strip())
    led = Counter()
    for l in lamps:
        led[l["settlement"]] += int(l["count"])
    order = {p["settlement"]: i for i, p in enumerate(places)}

    wb = Workbook()
    ws = wb.active
    ws.title = "Упатство"
    for line in [
        "Проверка на поврзувањето мерно место → населено место (улично осветлување, СН/НН LV1.1)",
        "",
        "1. Лист „Препознаени“: за секое мерно место во колоната „Точно?“ изберете Да или Не.",
        "   Ако е Не, во „Точно населено место“ изберете го точното место од листата (или „" + NOT_LIGHTING + "“).",
        "   Редовите со средна или ниска сигурност се обоени розово: нив проверете ги прво.",
        "2. Лист „Непрепознаени“: во „Точно населено место“ изберете каде припаѓа мерното место (жолтите ќелии).",
        "3. Лист „Населени места“: за секое од 79-те места е прикажано колку мерни места и LED светилки сме препознале.",
        "   Ако знаете мерни места што недостасуваат, запишете ги броевите во „Мерни места што недостасуваат“.",
        "4. Зачувајте ја табелата и испратете ја назад. Поправките се вчитуваат автоматски.",
        "",
        "Населено место = службено име од списокот на 79 населени места.",
        "Единица = населените места Горно/Долно и Големо/Мало се обединети по основното име; Гарани и Стрелци се заедно.",
    ]:
        ws.append([line])
    ws.column_dimensions["A"].width = 110
    ws["A1"].font = Font(bold=True, size=13)

    lists = wb.create_sheet("Листи")
    names = [p["settlement"] for p in places] + [NOT_LIGHTING]
    for i, n in enumerate(names, 1):
        lists.cell(i, 1, n)
    lists.cell(1, 2, "Да"); lists.cell(2, 2, "Не")
    lists.sheet_state = "hidden"
    dv_place = lambda: DataValidation(type="list", formula1=f"=Листи!$A$1:$A${len(names)}", allow_blank=True)
    dv_yes = DataValidation(type="list", formula1="=Листи!$B$1:$B$2", allow_blank=True)

    ok = wb.create_sheet("Препознаени", 1)
    header(ok, ["Поранешна општина", "Населено место", "Единица", "Мерно место", "Адреса (последна фактура)", "Други записи на адресата", "Сигурност", "kWh 12 мес.",
                "Точно? (Да/Не)", "Точно населено место (ако е грешка)", "Забелешка"], [14, 22, 20, 17, 40, 34, 11, 11, 10, 26, 30])
    known = sorted([r for r in mapping if r["settlement"]], key=lambda r: (GROUPS.index(r["group"]), order[r["settlement"]], r["address"]))
    for r in known:
        other = " / ".join(sorted(addresses[r["meter"]] - {r["address"]}))
        ok.append([r["group"], r["settlement"], r["unit"], r["meter"], r["address"], other, r["certainty"], round(kwh[r["meter"]]), None, None, None])
        if r["certainty"] != "висока":
            for c in ok[ok.max_row][:8]:
                c.fill = CHECK
    d1, d2 = dv_yes, dv_place()
    ok.add_data_validation(d1); ok.add_data_validation(d2)
    d1.add(f"I2:I{ok.max_row}"); d2.add(f"J2:J{ok.max_row}")
    ok.auto_filter.ref = f"A1:K{ok.max_row}"

    un = wb.create_sheet("Непрепознаени", 2)
    header(un, ["Поранешна општина (според бројот)", "Мерно место", "Адреса (последна фактура)", "Други записи на адресата", "kWh 12 мес.", "Точно населено место", "Забелешка"], [18, 17, 40, 34, 11, 26, 30])
    for r in sorted([r for r in mapping if not r["settlement"]], key=lambda r: -kwh[r["meter"]]):
        other = " / ".join(sorted(addresses[r["meter"]] - {r["address"]}))
        un.append([r["group"] or group_of_meter(r["meter"]), r["meter"], r["address"], other, round(kwh[r["meter"]]), None, None])
        un.cell(un.max_row, 6).fill = TODO
    d3 = dv_place(); un.add_data_validation(d3); d3.add(f"F2:F{max(un.max_row, 2)}")

    pl = wb.create_sheet("Населени места", 3)
    header(pl, ["Поранешна општина", "Населено место", "Единица", "Површина km²", "Препознаени мерни места", "Мерни места (броеви)", "LED светилки (ЈПП)",
                "Мерни места што недостасуваат", "Забелешка"], [14, 22, 20, 11, 11, 40, 11, 30, 30])
    by_place = defaultdict(list)
    for r in known:
        by_place[r["settlement"]].append(r["meter"])
    for p in places:
        ms = by_place[p["settlement"]]
        pl.append([p["group"], p["settlement"], p["unit"], float(p["area_km2"]), len(ms), ", ".join(ms), led[p["settlement"]], None, None])
        pl.cell(pl.max_row, 8).fill = TODO
    pl.auto_filter.ref = f"A1:I{pl.max_row}"

    IZLEZ.mkdir(exist_ok=True)
    wb.save(IZLEZ / "Проверка_мерни_места.xlsx")
    print("Препознаени:", len(known), " Непрепознаени:", un.max_row - 1, " Населени места:", len(places))


def uvezi(path):
    places = {p["settlement"]: p for p in read("naseleni.csv")}
    mapping = {r["meter"]: r for r in read("mapiranje.csv")}
    wb = load_workbook(path, data_only=True)
    changes = 0

    def assign(meter, place, how):
        nonlocal changes
        r = mapping[str(meter)]
        if place == NOT_LIGHTING:
            r.update(settlement="", unit="", certainty="не е улично осветлување")
        else:
            if place not in places:
                raise SystemExit(f"Непознато населено место „{place}“ за мерно место {meter}")
            r.update(settlement=place, unit=places[place]["unit"], group=places[place]["group"], certainty=how)
        changes += 1

    for row in wb["Препознаени"].iter_rows(min_row=2, values_only=True):
        meter, yes, place = row[3], row[8], row[9]
        if place:
            assign(meter, place, "поправено од корисник")
        elif yes == "Да":
            mapping[str(meter)]["certainty"] = "потврдено од корисник"; changes += 1
    for row in wb["Непрепознаени"].iter_rows(min_row=2, values_only=True):
        if row[5]:
            assign(row[1], row[5], "поправено од корисник")
    with open(VLEZ / "mapiranje.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS)
        w.writeheader(); w.writerows(mapping.values())
    print("Вчитани промени:", changes)


if __name__ == "__main__":
    if sys.argv[1:2] == ["izvezi"]:
        izvezi()
    elif sys.argv[1:2] == ["uvezi"] and len(sys.argv) == 3:
        uvezi(sys.argv[2])
    else:
        raise SystemExit(__doc__)
