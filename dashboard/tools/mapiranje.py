"""Ги поврзува мерните места за улично осветлување (СН/НН LV1.1) со
населените места од списокот на поставени LED светилки.

Излез:
  data/svetilki.csv   – списокот на светилки (населено место, производител, W, број)
  data/mapiranje.csv  – мерно место → населено место, со степен на сигурност

Употреба:  python3 dashboard/tools/mapiranje.py <список_светилки.xlsx>
"""
import csv
import re
import sys
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"

GROUPS = {
    "ОСЛОМЕЈ": "Осломеј", "ЗАЈАС": "Зајас", "ВРАНЕШТИЦА": "Вранештица",
    "ДРУГОВО": "Другово", "КИЧЕВО (Селата)": "Кичево – села", "КИЧЕВО (Градот)": "Кичево – град",
}


def read_lamps(path):
    ws = openpyxl.load_workbook(path, data_only=True).active
    out, group = [], None
    for row in ws.iter_rows(values_only=True):
        a, b, c, d, e, f = (list(row) + [None] * 6)[:6]
        if isinstance(b, str) and ("Општина" in b) and c is None:
            group = next((v for k, v in GROUPS.items() if k in b), None)
            continue
        if group and isinstance(a, int) and c == "ЛЕД Светилка" and isinstance(e, (int, float)):
            name = re.sub(r"[\s :]+$", "", re.sub(r"^[\s ]+", "", b))
            if group == "Кичево – град":
                name = "Кичево (град)"
            out.append({"group": group, "settlement": name, "street": b.strip(" : ") if group == "Кичево – град" else "",
                        "manufacturer": d, "watt": e, "count": f})
    return out


# (регуларен израз врз адресата, населено место, сигурност)
RULES = [
    # Осломеј
    (r"Ј\.ДОЛЕНЦИ", "Доленци", "висока"),
    (r"АРАН[ЃГ]ЕЛ", "Аранѓел", "висока"),
    (r"БЕРИКОВО", "Бериково", "висока"),
    (r"ЖУБРИНО", "Журбино", "висока"),
    (r"Н\.\s?СЕЛО|НОВО СЕЛО", "Ново Село", "висока"),
    (r"ОСЛОМЕЈ ОСЛОМЕЈ", "Осломеј", "висока"),
    (r"ПАПРАДИШТА", "Папрадишта", "висока"),
    (r"ПОПОЈАНИ", "Попојани", "висока"),
    (r"ПРЕМКА", "Премка", "висока"),
    (r"ЈАГОЛ", "Јагол", "висока"),
    (r"ГАРАНИ|СТРЕЛЦИ", "Гарани-Стрелци", "висока"),
    (r"МАМУДОВЦИ", "Мамудовци", "висока"),
    (r"СРБИЦА", "Србица", "висока"),
    (r"ТУИН", "Туин", "висока"),
    (r"ЦРВ[ИЕ]ВЦИ", "Црвивци", "висока"),
    (r"ШУТОВО", "Шутово", "висока"),
    # Зајас
    (r"БАЧИШТА", "Бачишта", "висока"),
    (r"БУКОЈЧАНИ", "Букојчани", "висока"),
    (r"Г\.\s?СТРО", "Горна Строгомишта", "висока"),
    (r"Д\.\s?СТРОГ", "Долна Строгомишта", "висока"),
    (r"ДЛАПКИН\s?ДОЛ|ДЛ\.\s?ДОЛ|ЗАЈАС Д\.\s?ДОЛ", "Длапкин Дол", "средна"),
    (r"ГРЕШНИЦА|МУЧОСКО", "Грешница", "висока"),
    (r"КОЛ\.\s?БРАЈОВЦИ|КОЛАРИ", "Колари", "висока"),
    (r"КОЛИБАРИ|КОЛБАРИ", "Колибари", "средна"),
    (r"ЛЕШНИЦА", "Лешница", "висока"),
    (r"МИДИНЦИ", "Мидинци", "висока"),
    (r"ЗАЈАС С\.\s?РЕЧАНИ", "Речани", "висока"),
    (r"ТАЈМИШТ", "Тајмиште", "висока"),
    (r"ЗАЈАС (ЗАЈАС|ЗАЈАЗ|С\.\s?ЗАЈАС|Б\.П\.ТИВОЛИ|ТС ТЕ[ЌЧ]Е)", "Зајас", "висока"),
    # Вранештица
    (r"АТИШТА", "Атишта", "висока"),
    (r"КАРБУНИЦА", "Карбуница", "висока"),
    (r"СТАРОЕЦ", "Староец", "висока"),
    (r"ЧЕЛОПЕЦИ", "Челопеци", "висока"),
    (r"ВРАНЕШТИЦА С\.\s?ВРАНЕШ", "Вранештица", "висока"),
    # Другово
    (r"ЈУДОВО", "Јудово", "висока"),
    (r"БЕЛИЦА", "Белица", "висока"),
    (r"БРЖДАНИ", "Брждани", "висока"),
    (r"ВИДРАНИ", "Видрани", "висока"),
    (r"Г\.\s?ДОБРЕНОЕЦ", "Горни Доброноец", "висока"),
    (r"Д\.\s?ДОБРЕНОЕЦ", "Долни Доброноец", "висока"),
    (r"Г\.\s?ЦРСКО", "Големо Црско", "средна"),
    (r"ДРУГОВО (ДРУГОВО|С\.\s?ДРУГОВО)", "Другово", "висока"),
    (r"ЕХЛОВЕЦ", "Ехловец", "висока"),
    (r"ИВАНЧИШТА", "Иванчишта", "висока"),
    (r"ИЗВОР", "Извор", "висока"),
    (r"КЛАДНИК", "Кладник", "висока"),
    (r"КЛЕНОЕЦ", "Кленоец", "висока"),
    (r"КОЗИЦА", "Козица", "висока"),
    (r"ЛАВЧАНИ", "Лавчани", "висока"),
    (r"МАЛКОЕЦ", "Малкоец", "висока"),
    (r"ПОДВИС", "Подвис", "висока"),
    (r"МАН\.\s?СВ\.\s?ПРЕЧИСТА", "Манастирско", "ниска"),
    (r"ПОПОЕЦ", "Попоец", "висока"),
    (r"ПОПОЛ?ЖАНИ", "Поположани", "висока"),
    (r"ПРОСРАЊЕ|ПРОСТАЊЕ", "Простање", "средна"),
    (r"СВИЊИШТА", "Свињишта", "висока"),
    (r"СРБЈАНИ", "Србјани", "висока"),
    (r"ДРУГОВО ЦЕР", "Цер", "висока"),
    # Кичево – села
    (r"КНЕЖИНО", "Кнежино", "средна"),
    (r"ЛАЗ[ОА]РОВЦИ", "Лазаровци", "висока"),
    (r"ОСОЈ", "Осој", "висока"),
    (r"ТРАПЧИН ДОЛ", "Трапчин Дол", "висока"),
]
# Адреси во Кичево што се села или други подрачја, а не градот
NOT_CITY = r"РАШТАНИ|БИЧИНЦИ|ДУШОГОБЕЦ|БЕТ\. СТОЛБ"


def settlement_of(addr):
    a = re.sub(r"\s+", " ", addr.upper())
    for pat, name, conf in RULES:
        if re.search(pat, a):
            return name, conf
    if a.startswith("КИЧЕВО") and not re.search(NOT_CITY, a):
        return "Кичево (град)", "висока"
    return "", "нема во списокот"


def main(xlsx):
    lamps = read_lamps(xlsx)
    with open(DATA / "svetilki.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["group", "settlement", "street", "manufacturer", "watt", "count"])
        w.writeheader(); w.writerows(lamps)
    known = {l["settlement"] for l in lamps}
    group_of = {l["settlement"]: l["group"] for l in lamps}

    last = {}
    for r in csv.DictReader(open(DATA / "stavki.csv", encoding="utf-8")):
        if "LV1" in r["category"]:
            last[r["meter"]] = r["address"].strip()
    rows = []
    for meter, addr in sorted(last.items(), key=lambda x: x[1]):
        name, conf = settlement_of(addr)
        if name and name not in known:
            raise SystemExit(f"Непознато населено место во правилата: {name}")
        rows.append({"meter": meter, "address": addr, "settlement": name, "group": group_of.get(name, ""), "certainty": conf})
    with open(DATA / "mapiranje.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["meter", "address", "settlement", "group", "certainty"])
        w.writeheader(); w.writerows(rows)
    print(len(lamps), "редови светилки,", sum(l["count"] for l in lamps), "светилки;",
          len(rows), "мерни места,", sum(1 for r in rows if r["settlement"]), "поврзани")


if __name__ == "__main__":
    main(sys.argv[1])
