"""Работна анализа на улично осветлување – верзија 3 (модел „часовник“).

Претпоставки (ваши): сите фактурирани kWh, освен новогодишното украсување на
наведените броила, ги трошат светилките; светилките светат само кога ќе им
дозволи часовникот; LED од ЈПП се 100 % исправни; стари светилки има според
вашата проценка по поранешна општина, со LED сијалица од 9 W, 80 % активни.

За секоја поранешна општина и месец се бара колку часа дневно светат:
  моќност (kW)   = LED од ЈПП + стари × 9 W × 0,8
  корисни kWh    = (фактурирани kWh − украсување) × (1 − загуби)
  часови на ден  = корисни kWh ÷ моќност ÷ денови во месецот
Се споредува со ноќта (зајдисонце–изгрејсонце за Кичево). Вишокот е можна
заштеда ако часовниците се наместат точно.

Обратна проверка: колку стари светилки би требале ако светилките светат точно
од зајдисонце до изгрејсонце (и до 2 часа подолго).

Употреба:  python3 analiza-svetilki/alatki/analiza.py
Излез:     analiza-svetilki/izlez/*.csv и Анализа_светилки.xlsx
"""
import calendar
import csv
import datetime as dt
import math
from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
VLEZ, IZLEZ = ROOT / "vlez", ROOT / "izlez"
GROUPS = ["Кичево", "Другово", "Осломеј", "Вранештица", "Зајас"]
MONTH_NAMES = {"Јануари": "01", "Февруари": "02", "Март": "03", "Април": "04", "Мај": "05", "Јуни": "06",
               "Јули": "07", "Август": "08", "Септември": "09", "Октомври": "10", "Ноември": "11", "Декември": "12"}
SHORT = {"01": "Јан", "02": "Фев", "03": "Мар", "04": "Апр", "05": "Мај", "06": "Јун",
         "07": "Јул", "08": "Авг", "09": "Сеп", "10": "Окт", "11": "Ное", "12": "Дек"}
# Месеци со новогодишно украсување (15.12–15.01) на броилата од vlez/ukrasi.csv
DECO_MONTHS = {"2025-12", "2026-01"}
EXTRA_MAX = 2  # горна граница за обратната проверка: часови на ден подолго од ноќта
# Поранешна општина според почетокот на бројот на мерното место (за непрепознаените)
METER_PREFIX = {"21021": "Кичево", "210316": "Другово", "210365": "Вранештица", "210366": "Зајас", "210367": "Осломеј"}


def read(name):
    return list(csv.DictReader(open(VLEZ / name, encoding="utf-8")))


def month_key(m):
    n, y = m.split()
    return y + "-" + MONTH_NAMES[n]


def night_hours(months, lat=41.5127, lon=20.9589):
    """Часови од зајдисонце до изгрејсонце по месец (NOAA, CET/CEST)."""
    rad = math.pi / 180
    def last_sunday(y, m):
        d = (dt.date(y, m + 1, 1) if m < 12 else dt.date(y + 1, 1, 1)) - dt.timedelta(days=1)
        return d.day - (d.weekday() + 1) % 7
    out = defaultdict(float)
    for k in months:
        y, m = map(int, k.split("-"))
        d = dt.date(y, m, 1)
        while d.month == m:
            g = 2 * math.pi / (366 if y % 4 == 0 else 365) * (d.timetuple().tm_yday - 1)
            eq = 229.18 * (0.000075 + 0.001868 * math.cos(g) - 0.032077 * math.sin(g) - 0.014615 * math.cos(2 * g) - 0.040849 * math.sin(2 * g))
            dec = 0.006918 - 0.399912 * math.cos(g) + 0.070257 * math.sin(g) - 0.006758 * math.cos(2 * g) + 0.000907 * math.sin(2 * g) - 0.002697 * math.cos(3 * g) + 0.00148 * math.sin(3 * g)
            ha = math.acos(math.cos(90.833 * rad) / (math.cos(lat * rad) * math.cos(dec)) - math.tan(lat * rad) * math.tan(dec)) / rad
            dst = 3 < m < 10 or (m == 3 and d.day >= last_sunday(y, 3)) or (m == 10 and d.day < last_sunday(y, 10))
            tz = 2 if dst else 1
            rise, set_ = (720 - 4 * (lon + ha) - eq) / 60 + tz, (720 - 4 * (lon - ha) - eq) / 60 + tz
            out[k] += 24 - (set_ - rise)
            d += dt.timedelta(days=1)
    return dict(out)


def group_of_meter(meter):
    return next((g for p, g in METER_PREFIX.items() if meter.startswith(p)), "")


def analyse():
    a = {r["kluch"]: r["vrednost"] for r in read("pretpostavki.csv")}
    W_OLD, ACTIVE, LOSS = float(a["stari_watt"]), float(a["stari_aktivni"]), float(a["zagubi"])
    places, lamps, mapping = read("naseleni.csv"), read("svetilki.csv"), read("mapiranje.csv")
    deco_meters = {r["meter"] for r in read("ukrasi.csv")}
    shares = {r["opstina"]: r for r in read("procenki.csv")}

    units = {}
    for p in places:
        u = units.setdefault(p["unit"], {"group": p["group"], "members": [], "led": 0, "led_w": 0.0, "meters": set()})
        u["members"].append(p["settlement"])
    for g in GROUPS:
        units["Непознато место – " + g] = {"group": g, "members": [], "led": 0, "led_w": 0.0, "meters": set()}
    for l in lamps:
        units[l["unit"]]["led"] += int(l["count"])
        units[l["unit"]]["led_w"] += int(l["count"]) * float(l["watt"])
    unit_of, meta = {}, {r["meter"]: r for r in mapping}
    for r in mapping:
        if r["certainty"] == "не е улично осветлување":
            continue
        unit_of[r["meter"]] = r["unit"] or "Непознато место – " + (r["group"] or group_of_meter(r["meter"]))

    per_meter, billed = defaultdict(dict), defaultdict(dict)
    months = set()
    for r in read("stavki.csv"):
        if "LV1" not in r["category"] or r["meter"] not in unit_of:
            continue
        m = month_key(r["month"])
        months.add(m)
        per_meter[r["meter"]][m] = per_meter[r["meter"]].get(m, 0) + float(r["kwh"])
        billed[r["meter"]][m] = billed[r["meter"]].get(m, 0) + float(r["billed"])
        units[unit_of[r["meter"]]]["meters"].add(r["meter"])
    months = sorted(months)
    H = night_hours(months)
    D = {m: calendar.monthrange(int(m[:4]), int(m[5:]))[1] for m in months}
    fit_months = [m for m in months if m not in DECO_MONTHS]

    def fit(v):
        """kWh_месец ≈ a × ноќни часови + b × денови (без месеците со украсување)."""
        ms = [m for m in fit_months if m in v]
        x1, x2, y = [H[m] for m in ms], [D[m] for m in ms], [v[m] for m in ms]
        if sum(1 for e in y if e > 0) < 3:
            return sum(v.values()) / sum(H[m] for m in v), 0.0
        s11, s22, s12 = sum(p * p for p in x1), sum(q * q for q in x2), sum(p * q for p, q in zip(x1, x2))
        t1, t2 = sum(p * e for p, e in zip(x1, y)), sum(q * e for q, e in zip(x2, y))
        det = s11 * s22 - s12 * s12
        na, nb = (t1 * s22 - t2 * s12) / det, (s11 * t2 - s12 * t1) / det
        if na < 0:
            return 0.0, t2 / s22
        if nb < 0:
            return t1 / s11, 0.0
        return na, nb

    comp = defaultdict(lambda: defaultdict(float))
    profiles = []
    for meter, v in per_meter.items():
        night_kw, per_day = fit(v)
        u = unit_of[meter]
        deco_sum = 0.0
        for m in months:
            e = v.get(m, 0.0)
            deco = max(0.0, e - night_kw * H[m] - per_day * D[m]) if meter in deco_meters and m in DECO_MONTHS else 0.0
            c = comp[u, m]
            c["kwh"] += e; c["deco"] += deco; c["billed"] += billed[meter].get(m, 0.0)
            deco_sum += deco
        year_night, year_flat = night_kw * sum(H.values()), per_day * sum(D.values())
        flat = year_flat / (year_night + year_flat) if year_night + year_flat else 0
        winter = sum(v.get(m, 0) / D[m] for m in ("2025-11", "2026-02"))
        summer = sum(v.get(m, 0) / D[m] for m in ("2025-07", "2026-06"))
        profiles.append({
            "meter": meter, "mesto": u, "opstina": units[u]["group"], "adresa": meta[meter]["address"], "sigurnost": meta[meter]["certainty"],
            "ukras": "да" if meter in deco_meters else "", "kwh_ukras": round(deco_sum, 1), "kwh": round(sum(v.values()), 1),
            "nokna_moknost_kw": round(night_kw, 3), "dneven_del_kwh": round(per_day, 2), "udel_ramen": round(flat, 3),
            "zima_leto": round(winter / summer, 2) if summer > 0 else None,
            "profil": "без потрошувачка" if not sum(v.values()) else "следи ноќта" if flat < 0.15 else "рамен" if flat > 0.6 else "мешан",
        })

    def unit_row(name, u, m, ks):
        c = {f: sum(comp[name, k][f] for k in ks) for f in ("kwh", "deco", "billed")}
        h, d = sum(H[k] for k in ks), sum(D[k] for k in ks)
        lamps_kwh = c["kwh"] - c["deco"]
        useful = lamps_kwh * (1 - LOSS)
        led_kw = u["led_w"] / 1000
        rest = useful - led_kw * h
        old_astro = max(rest, 0) * 1000 / (h * W_OLD) / ACTIVE
        deficit = max(-rest, 0) / (h * led_kw / u["led"]) if u["led"] else 0
        return {
            "opstina": u["group"], "mesto": name, "naseleni_mesta": ", ".join(u["members"]), "mesec": m,
            "merni_mesta": len(u["meters"]), "led": u["led"], "led_kw": round(led_kw, 3), "denovi": d,
            "nokni_casovi": round(h, 1), "nok_den": round(h / d, 2), "kwh": round(c["kwh"], 2), "kwh_ukras": round(c["deco"], 1),
            "kwh_zagubi": round(lamps_kwh * LOSS, 1), "kwh_korisno": round(useful, 1), "kwh_led_nok": round(led_kw * h, 1),
            "led_casovi_den": round(useful / led_kw / d, 2) if led_kw else None,
            "stari_postaveni": round(old_astro), "led_ne_svetat": round(deficit),
            "fakturirano_den": round(c["billed"]),
            "status": ("Нема мерно место" if not u["meters"] else "Без потрошувачка" if not c["kwh"] else
                       "Нема LED во списокот" if not u["led"] else "Помалку од LED цела ноќ" if rest < 0 else "Повеќе од LED цела ноќ"),
        }

    rows = []
    for name, u in units.items():
        if name.startswith("Непознато") and not u["meters"]:
            continue
        for m in months + ["all"]:
            rows.append(unit_row(name, u, m, months if m == "all" else [m]))

    def group_old(g):
        if g == "Вкупно":
            return sum(group_old(x) for x in GROUPS)
        led = sum(u["led"] for u in units.values() if u["group"] == g)
        s = float(shares[g]["udel_stari"])
        return led * s / (1 - s)

    by_group = []
    for g in GROUPS + ["Вкупно"]:
        old_n = group_old(g)
        for m in months + ["all"]:
            rs = [r for r in rows if r["mesec"] == m and (g == "Вкупно" or r["opstina"] == g)]
            h, d = rs[0]["nokni_casovi"], rs[0]["denovi"]
            led, led_kw = sum(r["led"] for r in rs), sum(r["led_kw"] for r in rs)
            kwh, deco, useful = (sum(r[f] for r in rs) for f in ("kwh", "kwh_ukras", "kwh_korisno"))
            fakt = sum(r["fakturirano_den"] for r in rs)
            p_kw = led_kw + old_n * W_OLD * ACTIVE / 1000
            hours = useful / p_kw
            save_useful = max(0.0, useful - p_kw * h)
            save_kwh = save_useful / (1 - LOSS)
            h2 = h + EXTRA_MAX * d
            by_group.append({
                "opstina": g, "mesec": m, "denovi": d, "led": led, "led_kw": round(led_kw, 2),
                "stari_procenka": round(old_n), "moknost_kw": round(p_kw, 2),
                "kwh": round(kwh, 1), "kwh_ukras": round(deco, 1), "kwh_zagubi": round((kwh - deco) * LOSS, 1), "kwh_korisno": round(useful, 1),
                "kwh_led_nok": round(led_kw * h, 1), "kwh_stari_nok": round(old_n * W_OLD * ACTIVE * h / 1000, 1),
                "kwh_vishok": round(useful - p_kw * h, 1),
                "nokni_casovi": round(h, 1), "nok_den": round(h / d, 2), "casovi": round(hours, 1), "casovi_den": round(hours / d, 2),
                "vishok_den": round(hours / d - h / d, 2),
                "ushteda_kwh": round(save_kwh), "ushteda_den": round(save_kwh * (fakt / kwh) if kwh else 0),
                "stari_potrebni_nok": round(max(0.0, useful - led_kw * h) * 1000 / (h * W_OLD) / ACTIVE),
                "stari_potrebni_nok2": round(max(0.0, useful - led_kw * h2) * 1000 / (h2 * W_OLD) / ACTIVE),
                "ocenka": "можно" if hours / d - h / d <= EXTRA_MAX + 0.5 else "невозможно со оваа моќност",
            })

    compare = []
    for g in GROUPS + ["Вкупно"]:
        y = next(r for r in by_group if r["opstina"] == g and r["mesec"] == "all")
        monthly = [r["vishok_den"] for r in by_group if r["opstina"] == g and r["mesec"] != "all"]
        compare.append({
            "opstina": g, "led": y["led"], "procenka_zapis": shares[g]["zapis"] if g in shares else "збир (ваша проценка: над " + a["stari_procenka_vkupno"] + ")",
            "procenka_udel": float(shares[g]["udel_stari"]) if g in shares else round(y["stari_procenka"] / (y["stari_procenka"] + y["led"]), 4),
            "stari_procenka": y["stari_procenka"], "moknost_kw": y["moknost_kw"], "casovi_den": y["casovi_den"], "nok_den": y["nok_den"],
            "vishok_den": y["vishok_den"], "vishok_min": min(monthly), "vishok_max": max(monthly),
            "stari_potrebni_nok": y["stari_potrebni_nok"], "stari_potrebni_nok2": y["stari_potrebni_nok2"],
            "ushteda_kwh": y["ushteda_kwh"], "ushteda_den": y["ushteda_den"], "ocenka": y["ocenka"],
        })
    return a, months, rows, by_group, compare, sorted(profiles, key=lambda p: (GROUPS.index(p["opstina"]), p["mesto"], -p["kwh"]))


def write_csv(name, rows):
    with open(IZLEZ / name, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader(); w.writerows(rows)


HEAD = Font(bold=True)
FILL = PatternFill("solid", fgColor="E8EEF2")


def sheet(wb, title, header, data, widths=None):
    ws = wb.create_sheet(title)
    ws.append(header)
    for c in ws[1]:
        c.font, c.fill, c.alignment = HEAD, FILL, Alignment(wrap_text=True, vertical="top")
    for row in data:
        ws.append(row)
    ws.freeze_panes = "B2"
    for i, w in enumerate(widths or [], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    return ws


def write_xlsx(a, months, rows, by_group, compare, profiles):
    wb = Workbook()
    ws = wb.active
    ws.title = "Претпоставки"
    lines = [
        ("Работна анализа на улично осветлување – Општина Кичево (јули 2025 – јуни 2026), верзија 3: модел „часовник“", None),
        ("Ориентациона анализа со импровизирани претпоставки. Се дополнува со проверените мерни места и населени места.", None),
        ("", None),
        ("Претпоставка", "Вредност"),
        ("LED светилки заменети со ЈПП (список), 100 % исправни", a["led_vkupno"]),
        ("Моќност на стара светилка со LED сијалица (W)", a["stari_watt"]),
        ("Удел на активни стари светилки", a["stari_aktivni"]),
        ("Ваша проценка за вкупно незаменети стари светилки (над)", a["stari_procenka_vkupno"]),
        ("Загуби по броилото (удел од фактурираните kWh)", a["zagubi"]),
        ("Новогодишно украсување (само броилата од листот „Украсување“)", a["ukrasuvanje"]),
        ("", None),
        ("Пресметка за секоја поранешна општина и месец", None),
        ("1. Сите фактурирани kWh, освен украсувањето, ги трошат светилките; светат само кога дозволува часовникот.", None),
        ("2. Корисни kWh = (kWh − украсување) × (1 − загуби).", None),
        ("3. Моќност = LED од ЈПП + стари светилки (ваша проценка по општина) × 9 W × удел активни.", None),
        ("4. Часови на ден = корисни kWh ÷ моќност ÷ денови. Се споредува со ноќта (зајдисонце–изгрејсонце за Кичево).", None),
        ("5. Вишокот над ноќта е можна заштеда ако часовниците се наместат точно (kWh и денари по фактурираната цена).", None),
        ("6. Обратна проверка: колку стари светилки би требале ако светилките светат точно ноќта, односно до 2 часа подолго.", None),
        ("7. „Невозможно со оваа моќност“: потребни се повеќе од 2,5 часа дневно над ноќта, па инсталираната моќност мора да е поголема.", None),
        ("", None),
        ("Ноќни часови по месец", None),
    ]
    num = lambda v: float(v) if isinstance(v, str) and v.replace(".", "", 1).isdigit() else v
    for l in lines:
        ws.append([l[0], num(l[1])])
    for m in months:
        r = next(r for r in rows if r["mesec"] == m)
        ws.append([m, r["nokni_casovi"]])
    ws.column_dimensions["A"].width = 110
    ws.column_dimensions["B"].width = 14
    ws["A1"].font = Font(bold=True, size=13)
    for cell in ("A4", "B4", "A12", "A21"):
        ws[cell].font = HEAD

    mh = [SHORT[m[5:]] + " " + m[2:4] for m in months]
    g = {(r["opstina"], r["mesec"]): r for r in by_group}
    G = GROUPS + ["Вкупно"]
    sheet(wb, "Часови по општина", ["Поранешна општина", "LED", "Стари (проценка)", "Моќност kW"] + ["Ч/ден " + h for h in mh] +
          ["Ч/ден (12 мес.)", "Ноќ ч/ден", "Вишок ч/ден", "Оценка"],
          [[x, g[x, "all"]["led"], g[x, "all"]["stari_procenka"], g[x, "all"]["moknost_kw"]] + [g[x, m]["casovi_den"] for m in months] +
           [g[x, "all"]["casovi_den"], g[x, "all"]["nok_den"], g[x, "all"]["vishok_den"], g[x, "all"]["ocenka"]] for x in G],
          [18, 8, 10, 10] + [8] * len(months) + [10, 9, 10, 26])
    sheet(wb, "Енергија по општина", ["Поранешна општина", "kWh 12 мес.", "Украсување", "Загуби", "Корисни kWh", "LED цела ноќ", "Стари цела ноќ",
                                     "Вишок (подолго од ноќта)", "Заштеда kWh", "Заштеда ден.", "Стари потребни (точно ноќ)", "Стари потребни (ноќ + 2 ч)"],
          [[x] + [g[x, "all"][k] for k in ("kwh", "kwh_ukras", "kwh_zagubi", "kwh_korisno", "kwh_led_nok", "kwh_stari_nok", "kwh_vishok",
                                          "ushteda_kwh", "ushteda_den", "stari_potrebni_nok", "stari_potrebni_nok2")] for x in G],
          [18, 12, 11, 11, 12, 12, 12, 13, 12, 12, 13, 13])
    sheet(wb, "Детално по општина", ["Поранешна општина", "Месец", "Денови", "Ноќни часови", "kWh", "Украсување", "Загуби", "Корисни kWh", "LED цела ноќ",
                                    "Стари цела ноќ", "Вишок kWh", "Часови на ден", "Ноќ ч/ден", "Вишок ч/ден", "Заштеда kWh", "Заштеда ден.", "Оценка"],
          [[r["opstina"], r["mesec"], r["denovi"], r["nokni_casovi"], r["kwh"], r["kwh_ukras"], r["kwh_zagubi"], r["kwh_korisno"], r["kwh_led_nok"],
            r["kwh_stari_nok"], r["kwh_vishok"], r["casovi_den"], r["nok_den"], r["vishok_den"], r["ushteda_kwh"], r["ushteda_den"], r["ocenka"]]
           for r in by_group if r["mesec"] != "all"], [16, 9, 7, 9, 11, 10, 10, 11, 11, 11, 11, 9, 9, 9, 11, 11, 24])

    u_all = {(r["mesto"], r["mesec"]): r for r in rows}
    names = sorted({r["mesto"] for r in rows}, key=lambda n: (GROUPS.index(u_all[n, "all"]["opstina"]), n.startswith("Непознато"), -u_all[n, "all"]["kwh"]))
    sheet(wb, "По место и месец", ["Поранешна општина", "Населено место", "Мерни места", "LED"] + ["Стари* " + h for h in mh] +
          ["Стари* (12 мес.)", "LED што не светат", "LED часови/ден", "Ноќ ч/ден", "kWh 12 мес.", "Украсување", "Статус", "Населени места"],
          [[u_all[n, "all"]["opstina"], n, u_all[n, "all"]["merni_mesta"], u_all[n, "all"]["led"]] + [u_all[n, m]["stari_postaveni"] for m in months] +
           [u_all[n, "all"][k] for k in ("stari_postaveni", "led_ne_svetat", "led_casovi_den", "nok_den", "kwh", "kwh_ukras", "status", "naseleni_mesta")]
           for n in names], [14, 26, 8, 8] + [8] * len(months) + [10, 10, 10, 9, 11, 10, 24, 36])
    ws2 = wb["По место и месец"]
    ws2.append([])
    ws2.append(["* Стари = стари светилки (9 W, 80 % активни) потребни ако светилките светат точно од зајдисонце до изгрејсонце. LED часови/ден = колку би светеле LED ако се единствени."])

    sheet(wb, "Мерни места – профил", ["Поранешна општина", "Населено место", "Мерно место", "Адреса", "Украсување", "kWh 12 мес.", "Ноќна моќност kW",
                                       "Дневен рамен дел kWh/ден", "Удел рамен дел", "Зима ÷ лето (kWh по ден)", "Профил"],
          [[p["opstina"], p["mesto"], p["meter"], p["adresa"], p["ukras"], p["kwh"], p["nokna_moknost_kw"], p["dneven_del_kwh"], p["udel_ramen"], p["zima_leto"], p["profil"]]
           for p in profiles], [14, 26, 17, 36, 9, 11, 11, 11, 10, 12, 14])
    wsp = wb["Мерни места – профил"]
    wsp.append([])
    wsp.append(["Профил: „следи ноќта“ – потрошувачката по ден расте зиме и опаѓа лете (астрономски часовник); „рамен“ – иста секој месец (часовник на фиксни часови, без летно/зимско време или процена на читање); „мешан“ – меѓу двете."])

    sheet(wb, "Споредба со проценки", ["Поранешна општина", "LED", "Ваша проценка", "Удел стари", "Стари (проценка)", "Моќност kW", "Часови на ден",
                                       "Ноќ ч/ден", "Вишок ч/ден", "Вишок најмалку (месец)", "Вишок најмногу (месец)", "Стари потребни (точно ноќ)",
                                       "Стари потребни (ноќ + 2 ч)", "Заштеда kWh/год.", "Заштеда ден./год.", "Оценка"],
          [[c["opstina"], c["led"], c["procenka_zapis"], c["procenka_udel"], c["stari_procenka"], c["moknost_kw"], c["casovi_den"], c["nok_den"],
            c["vishok_den"], c["vishok_min"], c["vishok_max"], c["stari_potrebni_nok"], c["stari_potrebni_nok2"], c["ushteda_kwh"], c["ushteda_den"], c["ocenka"]]
           for c in compare], [18, 8, 24, 9, 11, 10, 9, 9, 9, 11, 11, 13, 13, 12, 12, 26])
    prof = {p["meter"]: p for p in profiles}
    sheet(wb, "Украсување", ["Мерно место", "Адреса", "Забелешка", "kWh 12 мес.", "kWh украсување (дек.+јан.)"],
          [[d["meter"], d["adresa"], d["zabeleska"], prof.get(d["meter"], {}).get("kwh"), prof.get(d["meter"], {}).get("kwh_ukras")] for d in read("ukrasi.csv")],
          [17, 36, 60, 12, 14])
    for wsx in wb.worksheets[1:]:
        for col in wsx.iter_cols(min_row=2):
            for c in col:
                head = str(wsx.cell(1, c.column).value)
                if isinstance(c.value, float) and 0 <= c.value <= 1 and ("Удел" in head or "удел" in head):
                    c.number_format = "0%"
                elif isinstance(c.value, float) and ("ч/ден" in head.lower() or "Ч/ден" in head or "часови/ден" in head.lower() or "kW" in head and "kWh" not in head or "÷" in head or "Часови на ден" in head):
                    c.number_format = "0.0"
                elif isinstance(c.value, (int, float)):
                    c.number_format = "#,##0"
    wb.save(IZLEZ / "Анализа_светилки.xlsx")


def main():
    a, months, rows, by_group, compare, profiles = analyse()
    IZLEZ.mkdir(exist_ok=True)
    write_csv("po_mesto_mesec.csv", rows)
    write_csv("po_opstina_mesec.csv", by_group)
    write_csv("sporedba.csv", compare)
    write_csv("merni_mesta_profil.csv", profiles)
    write_xlsx(a, months, rows, by_group, compare, profiles)
    for c in compare:
        print(f"{c['opstina']:11} LED {c['led']:5} стари {c['stari_procenka']:5} P {c['moknost_kw']:6.1f} kW | {c['casovi_den']:5.1f} ч/ден (ноќ {c['nok_den']:.1f}, "
              f"вишок {c['vishok_den']:+.1f}; {c['vishok_min']:+.1f}…{c['vishok_max']:+.1f}) | потребни стари {c['stari_potrebni_nok']}/{c['stari_potrebni_nok2']} | "
              f"заштеда {c['ushteda_kwh']:,} kWh {c['ushteda_den']:,} ден. | {c['ocenka']}")


if __name__ == "__main__":
    main()
