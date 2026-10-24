"""Работна анализа: колку стари (незаменети) светилки светат во секое
населено место, по месец.

Модел за секое населено место и месец:
  E_LED   = вкупна моќност на LED светилките од ЈПП × ноќни часови × удел на работа
  остаток = фактурирани kWh − E_LED
  активни стари = остаток ÷ (ноќни часови × моќност на стара светилка)
  поставени стари = активни стари ÷ удел на активни
Ако остатокот е негативен, потрошувачката е помала од онаа што би ја имале
LED светилките, па дел од LED не светат (LED „дефицит“).

Вашите проценки за уделот на стари светилки по поранешна општина се користат
само на крајот, во листот „Споредба со проценки“.

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
# Поранешна општина според почетокот на бројот на мерното место (за непрепознаените)
# Новогодишно украсување во градот (15.12–15.01): тие месеци не влегуваат во
# поделбата ноќен/постојан товар, а вишокот над неа во градот е украсување.
DECO_MONTHS = {"2025-12", "2026-01"}
DECO_UNIT = "Кичево (град)"
H_YEAR = lambda H: sum(v for k, v in H.items() if k != "all")
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
    W_OLD, ACTIVE, BURN = float(a["stari_watt"]), float(a["stari_aktivni"]), float(a["led_rabota"])
    places, lamps, mapping = read("naseleni.csv"), read("svetilki.csv"), read("mapiranje.csv")

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

    per_meter = defaultdict(dict)
    months = set()
    for r in read("stavki.csv"):
        if "LV1" not in r["category"] or r["meter"] not in unit_of:
            continue
        m = month_key(r["month"])
        months.add(m)
        per_meter[r["meter"]][m] = per_meter[r["meter"]].get(m, 0) + float(r["kwh"])
        units[unit_of[r["meter"]]]["meters"].add(r["meter"])
    months = sorted(months)
    H = night_hours(months)
    D = {m: calendar.monthrange(int(m[:4]), int(m[5:]))[1] for m in months}
    fit_months = [m for m in months if m not in DECO_MONTHS]

    def fit(v):
        """kWh_месец ≈ ноќна моќност × ноќни часови + постојана моќност × 24 × денови (без месеците со украсување)."""
        ms = [m for m in fit_months if m in v]
        x1, x2, y = [H[m] for m in ms], [24 * D[m] for m in ms], [v[m] for m in ms]
        if sum(1 for e in y if e > 0) < 3:
            return sum(v.values()) / sum(H[m] for m in v), 0.0
        s11, s22, s12 = sum(a * a for a in x1), sum(b * b for b in x2), sum(a * b for a, b in zip(x1, x2))
        t1, t2 = sum(a * e for a, e in zip(x1, y)), sum(b * e for b, e in zip(x2, y))
        det = s11 * s22 - s12 * s12
        night, const = (t1 * s22 - t2 * s12) / det, (s11 * t2 - s12 * t1) / det
        if night < 0:
            return 0.0, t2 / s22
        if const < 0:
            return t1 / s11, 0.0
        return night, const

    comp = defaultdict(lambda: defaultdict(float))
    profiles = []
    for meter, v in per_meter.items():
        night_kw, const_kw = fit(v)
        u = unit_of[meter]
        for m in months:
            e = v.get(m, 0.0)
            const = min(e, const_kw * 24 * D[m])
            deco = max(0.0, e - const - night_kw * H[m]) if u == DECO_UNIT and m in DECO_MONTHS else 0.0
            c = comp[u, m]
            c["kwh"] += e; c["const"] += const; c["deco"] += deco; c["night"] += e - const - deco
        year_night, year_const = night_kw * sum(H[m] for m in months), const_kw * 24 * sum(D.values())
        share_const = year_const / (year_night + year_const) if year_night + year_const else 0
        winter = [v.get(m, 0) / D[m] for m in ("2025-11", "2026-02")]
        summer = [v.get(m, 0) / D[m] for m in ("2025-07", "2026-06")]
        ratio = (sum(winter) / sum(summer)) if sum(summer) > 0 else None
        profiles.append({
            "meter": meter, "mesto": u, "opstina": units[u]["group"], "adresa": meta[meter]["address"], "sigurnost": meta[meter]["certainty"],
            "kwh": round(sum(v.values()), 1),
            "nokna_moknost_kw": round(night_kw, 3), "postojana_moknost_kw": round(const_kw, 3),
            "udel_postojano": round(share_const, 3), "zima_leto": round(ratio, 2) if ratio else None,
            "profil": "без потрошувачка" if not sum(v.values()) else "ноќен" if share_const < 0.15 else "постојан" if share_const > 0.6 else "мешан",
        })

    def estimate(night, led, led_w, h):
        e_led = led_w * h * BURN / 1000
        rest = night - e_led
        active = max(rest, 0) * 1000 / (h * W_OLD)
        deficit = max(-rest, 0) * 1000 / (h * BURN * (led_w / led)) if led else 0
        return e_led, rest, active, active / ACTIVE, deficit

    def status(e, led, rest, n_meters):
        if not n_meters:
            return "Нема мерно место"
        if e == 0:
            return "Без потрошувачка"
        if not led:
            return "Само стари светилки"
        if rest < 0:
            return "LED не светат сите"
        return "LED и стари светилки"

    rows = []
    for name, u in units.items():
        if name.startswith("Непознато") and not u["meters"]:
            continue
        for m in months + ["all"]:
            ks = months if m == "all" else [m]
            c = {f: sum(comp[name, k][f] for k in ks) for f in ("kwh", "const", "deco", "night")}
            h = sum(H[k] for k in ks)
            e_led, rest, active, installed, deficit = estimate(c["night"], u["led"], u["led_w"], h)
            rows.append({
                "opstina": u["group"], "mesto": name, "naseleni_mesta": ", ".join(u["members"]), "mesec": m,
                "merni_mesta": len(u["meters"]), "led": u["led"], "led_kw": round(u["led_w"] / 1000, 3),
                "nokni_casovi": round(h, 1), "kwh": round(c["kwh"], 2), "kwh_postojano": round(c["const"], 1), "kwh_ukras": round(c["deco"], 1),
                "kwh_nokno": round(c["night"], 1), "kwh_led": round(e_led, 1), "ostatok_kwh": round(rest, 1),
                "stari_aktivni": round(active), "stari_postaveni": round(installed), "led_ne_svetat": round(deficit),
                "udel_stari": round(installed / (installed + u["led"]), 4) if installed + u["led"] else None,
                "status": status(c["kwh"], u["led"], rest, len(u["meters"])),
            })

    by_group = []
    for g in GROUPS + ["Вкупно"]:
        for m in months + ["all"]:
            rs = [r for r in rows if r["mesec"] == m and (g == "Вкупно" or r["opstina"] == g)]
            led = sum(r["led"] for r in rs)
            inst = sum(r["stari_postaveni"] for r in rs)
            h = sum(H[k] for k in (months if m == "all" else [m]))
            net = max(0.0, sum(r["kwh_nokno"] - r["kwh_led"] for r in rs)) * 1000 / (h * W_OLD) / ACTIVE
            by_group.append({
                "opstina": g, "mesec": m, "led": led, "kwh": round(sum(r["kwh"] for r in rs), 1),
                "kwh_postojano": round(sum(r["kwh_postojano"] for r in rs), 1), "kwh_ukras": round(sum(r["kwh_ukras"] for r in rs), 1),
                "kwh_nokno": round(sum(r["kwh_nokno"] for r in rs), 1), "kwh_led": round(sum(r["kwh_led"] for r in rs), 1),
                "ostatok_kwh": round(sum(max(r["ostatok_kwh"], 0) for r in rs), 1),
                "stari_aktivni": sum(r["stari_aktivni"] for r in rs), "stari_postaveni": inst,
                "led_ne_svetat": sum(r["led_ne_svetat"] for r in rs),
                "ostatok_neto": round(max(0.0, sum(r["kwh_nokno"] - r["kwh_led"] for r in rs)), 1),
                "stari_neto": round(net), "udel_stari": round(net / (net + led), 4) if net + led else None,
            })

    guesses = {r["opstina"]: r for r in read("procenki.csv")}
    compare = []
    for g in GROUPS + ["Вкупно"]:
        y = next(r for r in by_group if r["opstina"] == g and r["mesec"] == "all")
        monthly = [r["stari_neto"] for r in by_group if r["opstina"] == g and r["mesec"] != "all"]
        if g == "Вкупно":
            guess_n = sum(c["procenka_stari"] for c in compare)
            share = guess_n / (guess_n + y["led"])
            note = "збир; ваша вкупна проценка: над " + f"{float(a['stari_procenka_vkupno']):.0f}"
        else:
            share = float(guesses[g]["udel_stari"])
            guess_n = round(y["led"] * share / (1 - share))
            note = guesses[g]["zapis"]
        kwh_guess = guess_n * ACTIVE * W_OLD * H_YEAR(H) / 1000
        compare.append({
            "opstina": g, "led": y["led"], "presmetani_stari": y["stari_neto"], "stari_bruto": y["stari_postaveni"], "presmetan_udel": y["udel_stari"],
            "stari_min_mesec": min(monthly), "stari_max_mesec": max(monthly),
            "procenka_udel": round(share, 4), "procenka_zapis": note, "procenka_stari": guess_n,
            "kwh": y["kwh"], "kwh_postojano": y["kwh_postojano"], "kwh_ukras": y["kwh_ukras"],
            "ostatok_kwh": round(max(0.0, y["kwh_nokno"] - y["kwh_led"]), 1), "kwh_po_procenka": round(kwh_guess),
            "neobjasneti_kwh": round(max(0.0, y["kwh_nokno"] - y["kwh_led"]) - kwh_guess),
            "potrebna_moknost_w": round(max(0.0, y["kwh_nokno"] - y["kwh_led"]) * 1000 / (H_YEAR(H) * ACTIVE * guess_n), 1) if guess_n else None,
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
        ("Работна анализа на број на светилки – Општина Кичево (јули 2025 – јуни 2026)", None),
        ("Ова е ориентациона анализа со импровизирани претпоставки. Се дополнува со проверените мерни места и населени места.", None),
        ("", None),
        ("Претпоставка", "Вредност"),
        ("LED светилки заменети со ЈПП (список)", a["led_vkupno"]),
        ("LED светат од зајдисонце до изгрејсонце со полна моќност (удел од ноќта)", a["led_rabota"]),
        ("Моќност на стара светилка со LED сијалица (W)", a["stari_watt"]),
        ("Удел на активни стари светилки", a["stari_aktivni"]),
        ("Ваша проценка за вкупно незаменети стари светилки (над)", a["stari_procenka_vkupno"]),
        ("Новогодишно украсување на мерните места во градот", a["ukrasuvanje"]),
        ("", None),
        ("Пресметка (верзија 2: потрошувачката на секое мерно место се дели на три дела)", None),
        ("1. За секое мерно место: kWh по месец ≈ ноќна моќност × ноќни часови + постојана моќност × 24 ч × денови (без декември и јануари).", None),
        ("   Постојаниот дел троши и дење: светилки што светат дење или други потрошувачи на ормарот (згради, пумпи, бензински станици).", None),
        ("2. Украсување: во градот Кичево, вишокот во декември и јануари (15.12–15.01) над поделбата од чекор 1.", None),
        ("3. Ноќна енергија = фактурирани kWh − постојан дел − украсување. Само таа се споредува со светилките.", None),
        ("4. E_LED = вкупна моќност на LED светилките × ноќни часови (зајдисонце–изгрејсонце за Кичево); остаток = ноќна енергија − E_LED", None),
        ("5. Активни стари светилки = остаток ÷ (ноќни часови × моќност на стара светилка); поставени = активни ÷ удел на активни", None),
        ("6. Ако остатокот е негативен: дел од LED не светат („LED што не светат“).", None),
        ("   Нето по поранешна општина: вишокот и недостигот на сите места се сметаат заедно, па грешките во поврзувањето мерно место → село се поништуваат.", None),
        ("   Бруто: збир само на вишоците по места (поголем, зависи од точноста на поврзувањето).", None),
        ("7. Вашите проценки по поранешна општина се споредуваат само на крајот (лист „Споредба со проценки“).", None),
        ("", None),
        ("Ноќни часови по месец", None),
    ]
    num = lambda v: float(v) if isinstance(v, str) and v.replace(".", "", 1).isdigit() else v
    for l in lines:
        ws.append([l[0], num(l[1])])
    for m in months:
        h = next(r["nokni_casovi"] for r in rows if r["mesec"] == m)
        ws.append([m, h])
    ws.column_dimensions["A"].width = 95
    ws.column_dimensions["B"].width = 14
    ws["A1"].font = Font(bold=True, size=13)
    ws["A4"].font = ws["B4"].font = HEAD
    ws["A11"].font = ws["A24"].font = HEAD

    mh = [SHORT[m[5:]] + " " + m[2:4] for m in months]
    g_all = {(r["opstina"], r["mesec"]): r for r in by_group}
    sheet(wb, "По општина и месец", ["Поранешна општина", "LED (ЈПП)"] + ["Стари (нето) " + h for h in mh] + ["Стари нето (12 мес.)", "Удел стари", "Стари бруто (збир по места)",
          "LED што не светат (збир по места)", "kWh 12 мес.", "Постојан дел kWh", "Украсување kWh", "Ноќна енергија kWh", "kWh LED", "Остаток kWh (позитивен по места)"],
          [[g, g_all[g, "all"]["led"]] + [g_all[g, m]["stari_neto"] for m in months] +
           [g_all[g, "all"][k] for k in ("stari_neto", "udel_stari", "stari_postaveni", "led_ne_svetat", "kwh", "kwh_postojano", "kwh_ukras", "kwh_nokno", "kwh_led", "ostatok_kwh")]
           for g in GROUPS + ["Вкупно"]], [20, 10] + [9] * len(months) + [12, 10, 12, 14, 12, 12, 12, 12, 12, 14])

    u_all = {(r["mesto"], r["mesec"]): r for r in rows}
    names = sorted({r["mesto"] for r in rows}, key=lambda n: (GROUPS.index(u_all[n, "all"]["opstina"]), n.startswith("Непознато"), -u_all[n, "all"]["kwh"]))
    sheet(wb, "По место и месец", ["Поранешна општина", "Населено место", "Мерни места", "LED (ЈПП)"] + ["Стари " + h for h in mh] +
          ["Стари (12 мес.)", "Удел стари", "LED што не светат", "kWh 12 мес.", "Постојан дел kWh", "Украсување kWh", "Ноќна енергија kWh", "kWh LED", "Остаток kWh", "Статус", "Населени места во единицата"],
          [[u_all[n, "all"]["opstina"], n, u_all[n, "all"]["merni_mesta"], u_all[n, "all"]["led"]] + [u_all[n, m]["stari_postaveni"] for m in months] +
           [u_all[n, "all"][k] for k in ("stari_postaveni", "udel_stari", "led_ne_svetat", "kwh", "kwh_postojano", "kwh_ukras", "kwh_nokno", "kwh_led", "ostatok_kwh", "status", "naseleni_mesta")]
           for n in names], [14, 26, 8, 8] + [8] * len(months) + [10, 9, 10, 11, 11, 11, 11, 11, 11, 22, 36])

    sheet(wb, "Детално по месец", ["Поранешна општина", "Населено место", "Месец", "Мерни места", "LED", "LED kW", "Ноќни часови", "kWh", "Постојан дел kWh", "Украсување kWh",
                                   "Ноќна енергија kWh", "kWh LED", "Остаток kWh", "Активни стари", "Поставени стари", "LED што не светат", "Удел стари", "Статус"],
          [[r["opstina"], r["mesto"], r["mesec"], r["merni_mesta"], r["led"], r["led_kw"], r["nokni_casovi"], r["kwh"], r["kwh_postojano"], r["kwh_ukras"], r["kwh_nokno"],
            r["kwh_led"], r["ostatok_kwh"], r["stari_aktivni"], r["stari_postaveni"], r["led_ne_svetat"], r["udel_stari"], r["status"]] for r in rows if r["mesec"] != "all"],
          [14, 26, 9, 8, 7, 8, 9, 10, 10, 10, 10, 10, 10, 9, 9, 9, 9, 22])

    sheet(wb, "Мерни места – профил", ["Поранешна општина", "Населено место", "Мерно место", "Адреса", "kWh 12 мес.", "Ноќна моќност kW", "Постојана моќност kW",
                                       "Удел постојан дел", "Зима ÷ лето (kWh по ден)", "Профил"],
          [[p["opstina"], p["mesto"], p["meter"], p["adresa"], p["kwh"], p["nokna_moknost_kw"], p["postojana_moknost_kw"], p["udel_postojano"], p["zima_leto"], p["profil"]] for p in profiles],
          [14, 26, 17, 36, 11, 11, 11, 11, 13, 14])

    sheet(wb, "Споредба со проценки", ["Поранешна општина", "LED (ЈПП)", "Пресметани стари, нето (12 мес.)", "Најмалку во месец", "Најмногу во месец", "Пресметан удел стари",
                                       "Ваша проценка (удел)", "Запис", "Стари според проценката", "Ноќна енергија над LED kWh", "kWh за стари според проценката", "Необјаснети kWh", "Потребна моќност по стара светилка (W)"],
          [[c["opstina"], c["led"], c["presmetani_stari"], c["stari_min_mesec"], c["stari_max_mesec"], c["presmetan_udel"], c["procenka_udel"], c["procenka_zapis"],
            c["procenka_stari"], c["ostatok_kwh"], c["kwh_po_procenka"], c["neobjasneti_kwh"], c["potrebna_moknost_w"]] for c in compare],
          [18, 10, 14, 11, 11, 12, 12, 26, 13, 12, 14, 12, 18])
    for wsx in wb.worksheets[1:]:
        for col in wsx.iter_cols(min_row=2):
            for c in col:
                head = str(wsx.cell(1, c.column).value)
                if isinstance(c.value, float) and 0 <= c.value <= 1 and ("Удел" in head or "удел" in head):
                    c.number_format = "0%"
                elif isinstance(c.value, float) and ("kW" in head and "kWh" not in head or "÷" in head):
                    c.number_format = "0.00"
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
        print(f"{c['opstina']:11} LED {c['led']:5}  стари {c['presmetani_stari']:6} ({c['stari_min_mesec']}–{c['stari_max_mesec']})  удел {c['presmetan_udel']:.0%}  "
              f"| проценка {c['procenka_udel']:.0%} → {c['procenka_stari']:5}  потребни W {c['potrebna_moknost_w']}")


if __name__ == "__main__":
    main()
