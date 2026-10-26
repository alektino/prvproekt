"""Анализа на стари светилки – верзија 5 (LED моќност според табелата на моќности).

Наод од контролните села: LED светилките не се секаде со моќноста од пописот
(вградувани „што имало“). Бројот на LED е како во пописот (6.500), а моќноста
се менува со замени во табелата на моќности 12,1 / 20,9 / 25,9 / 32,1 / 64,5 W:

  села, удел p ≥ 0:  дел p од светилките на пописот се всушност 64,5 W
  села, удел p < 0:  дел |p| се за еден степен послаби (25,9 → 20,9, 32,1 → 25,9 ...)
  град, удел q:      дел q од светилките од 64,5 W се всушност 32,1 W

Стари светилки по место = (измерена моќност − LED моќност) ÷ (10 W × 0,8), најмалку 0.
Измерена моќност = медијана на месечните (kWh − украсување) × 0,9 ÷ ноќни часови
(неправилните или проценетите читања не влијаат).

Удел p по поранешна општина се одредува така што бројот на стари да е еднаков на
проценката (Зајас 10 %, Осломеј 10 %, Кичево 20 %, Другово 40 %, Вранештица 70 % од
сите светилки); за градот q така што градот да има 550 стари (500–600). Ако
проценката е под најмалиот можен број (сите LED 64,5 W), се зема најмалиот број.
Проверки: контролните села Староец и Видрани (не се користат во одредувањето) и
независна пресметка со удел p одреден само од Староец. Бројот 1.500 е само репер.

Употреба:  python3 analiza-svetilki/alatki/analiza_v5.py
Излез:     izlez/v5_*.csv и izlez/Анализа_светилки_v5.xlsx
"""
import csv
import statistics
from collections import Counter, defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import analiza as A
import analiza_v4 as V

W_EFF = V.W_OLD * V.ACTIVE / 1000
TYPES = [12.1, 20.9, 25.9, 32.1, 64.5]
STRONG, CITY_DOWN = 64.5, 32.1
WEAKER = {12.1: 12.1, 20.9: 12.1, 25.9: 20.9, 32.1: 25.9, 64.5: 32.1}
CITY = "Кичево (град)"
CITY_TARGET, CITY_RANGE = 550, (500, 600)
COUNTED = {"Староец": (29, 26), "Видрани": (7, 23)}
REFERENCE = 1500


def led_kw(types, p):
    """LED kW на место со состав {моќност: број} при удел замени p."""
    if p >= 0:
        return sum(n * ((1 - p) * w + p * STRONG) for w, n in types.items()) / 1000
    return sum(n * ((1 + p) * w - p * WEAKER[w]) for w, n in types.items()) / 1000


def city_kw(types, q):
    return sum(n * (w if w != STRONG else (1 - q) * STRONG + q * CITY_DOWN) for w, n in types.items()) / 1000


def solve(f, target, lo, hi):
    """f опаѓа; враќа x во [lo, hi] со f(x) ≈ target (или граница)."""
    if f(hi) >= target:
        return hi
    if f(lo) <= target:
        return lo
    for _ in range(60):
        mid = (lo + hi) / 2
        lo, hi = (mid, hi) if f(mid) > target else (lo, mid)
    return (lo + hi) / 2


def load():
    months, H, monthly, annual, groups, scenarios, deco_rows, checks = V.analyse()
    types = defaultdict(Counter)
    for l in A.read("svetilki.csv"):
        types[l["unit"]][float(l["watt"])] += int(l["count"])
    useful = defaultdict(dict)
    for r in monthly:
        if r["kwh"] > 0:
            useful[r["mesto"]][r["mesec"]] = r["korisni_kwh"] / r["nokni_casovi"]
    hy = sum(H.values())
    places = []
    for r in annual:
        if not r["merni_mesta"]:
            continue
        places.append({"opstina": r["opstina"], "mesto": r["mesto"], "led": r["led"], "types": types[r["mesto"]],
                       "kw": statistics.median(useful[r["mesto"]].values()), "kw_god": r["korisni_kwh"] / hy, "meseci": len(useful[r["mesto"]])})
    share = {r["opstina"]: float(r["udel_stari"]) for r in A.read("procenki.csv")}
    led_g = Counter()
    for r in annual:
        led_g[r["opstina"]] += r["led"]
    targets = {g: led_g[g] * share[g] / (1 - share[g]) for g in A.GROUPS}
    return months, H, annual, places, share, targets


def old(pl, p, q, w=W_EFF, kw="kw"):
    led = city_kw(pl["types"], q) if pl["mesto"] == CITY else led_kw(pl["types"], p)
    return max(0.0, (pl[kw] - led) / w)


def estimate(places, p_by_group, q, w=W_EFF, kw="kw"):
    return {pl["mesto"]: old(pl, p_by_group[pl["opstina"]], q, w, kw) for pl in places}


def calibrate(places, targets, w=W_EFF, kw="kw", city_target=CITY_TARGET, p_max=1.0):
    city = next(pl for pl in places if pl["mesto"] == CITY)
    q = solve(lambda x: -old(city, 0, x, w, kw), -city_target, 0.0, 1.0)  # со q градот има повеќе стари
    p, flags = {}, {}
    for g in A.GROUPS:
        rs = [pl for pl in places if pl["opstina"] == g]
        fixed = old(city, 0, q, w, kw) if g == "Кичево" else 0.0
        total = lambda x: fixed + sum(old(pl, x, q, w, kw) for pl in rs if pl["mesto"] != CITY)
        p[g] = solve(total, targets[g], -1.0, p_max)
        flags[g] = "под најмалиот можен број" if total(p_max) > targets[g] + 0.5 else "над најголемиот можен број" if total(-1.0) < targets[g] - 0.5 else ""
    return p, q, flags


def control(places, p, q):
    """Контролните села: проценка од моделот наспроти избројаното, и замени што ги објаснуваат."""
    rows = []
    for u, (n_led, n_old) in COUNTED.items():
        pl = next(x for x in places if x["mesto"] == u)
        need = pl["kw"] - n_old * W_EFF
        strong = (need * 1000 - n_led * 25.9) / (STRONG - 25.9)
        fits = 0 <= strong <= n_led
        rows.append({"mesto": u, "led_izbrojani": n_led, "stari_izbrojani": n_old, "led_popis": pl["led"], "izmereno_kw": round(pl["kw"], 3),
                     "led_kw_popis": round(n_led * 25.9 / 1000, 3), "potrebno_led_kw": round(need, 3),
                     "zameni_64_5": round(strong) if fits else None,
                     "zakluchok": f"Се објаснува: {round(strong)} од {n_led} LED се 64,5 W" if fits else
                     f"Не се објаснува ни ако сите {n_led} LED се 64,5 W (недостасуваат {need - n_led * STRONG / 1000:.2f} kW)",
                     "v4_stari": round((pl["kw"] - led_kw(pl["types"], 0)) / W_EFF),
                     "v5_stari": round(old(pl, p[pl["opstina"]], q))})
    return rows


def analyse():
    months, H, annual, places, share, targets = load()
    p, q, flags = calibrate(places, targets)
    est = estimate(places, p, q)

    # Независна проверка: удел p само од Староец (избројано), применет на сите села.
    st = next(pl for pl in places if pl["mesto"] == "Староец")
    p_st = solve(lambda x: (st["kw"] - led_kw({25.9: COUNTED["Староец"][0]}, x)) / W_EFF, COUNTED["Староец"][1], -1.0, 1.0)

    def total(pv, qv, **kw):
        e = estimate(places, pv, qv, **kw)
        out = Counter()
        for pl in places:
            out[pl["opstina"]] += e[pl["mesto"]]
        out["Вкупно"] = sum(out[g] for g in A.GROUPS)
        out["град"] = e[CITY]
        return out

    scen = [("v4: LED како во пописот", {g: 0.0 for g in A.GROUPS}, 0.0)]
    scen.append(("v5 главна: замени според проценките, град 550", p, q))
    for c in CITY_RANGE:
        pc, qc, _ = calibrate(places, targets, city_target=c)
        scen.append((f"v5, град {c} стари", pc, qc))
    scen.append((f"Независно: удел од Староец ({p_st:.0%} од LED се 64,5 W), град без замени", {g: p_st for g in A.GROUPS}, 0.0))
    scen.append((f"Независно: удел од Староец, град {CITY_TARGET}", {g: (p_st if g != "Кичево" else p["Кичево"]) for g in A.GROUPS}, q))
    pc, qc, _ = calibrate(places, targets, p_max=p_st)
    scen.append((f"v5, удел замени најмногу колку во Староец ({p_st:.0%})", pc, qc))
    scen.append(("Најмалку можно: сите LED во селата 64,5 W, град 550", {g: 1.0 for g in A.GROUPS}, q))
    scenarios = []
    for name, pv, qv in scen:
        t = total(pv, qv)
        scenarios.append({"scenario": name, **{g: round(t[g]) for g in A.GROUPS + ["Вкупно"]}, "grad": round(t["град"])})
    for name, kw in (("v5 главна, стари 9 W × 80 %", {"w": 9 * V.ACTIVE / 1000}), ("v5 главна, стари 10 W × 90 %", {"w": 10 * 0.9 / 1000}),
                     ("v5 главна, годишна енергија наместо медијана", {"kw": "kw_god"})):
        pc, qc, _ = calibrate(places, targets, **kw)
        t = total(pc, qc, **kw)
        scenarios.append({"scenario": name, **{g: round(t[g]) for g in A.GROUPS + ["Вкупно"]}, "grad": round(t["град"])})

    rows = []
    for pl in places:
        g = pl["opstina"]
        led0 = led_kw(pl["types"], 0)
        led5 = city_kw(pl["types"], q) if pl["mesto"] == CITY else led_kw(pl["types"], p[g])
        nmin = old(pl, 1.0, 1.0) if pl["mesto"] != CITY else None
        rows.append({"opstina": g, "mesto": pl["mesto"], "led": pl["led"], "led_kw_popis": round(led0, 2), "led_kw_v5": round(led5, 2),
                     "izmereno_kw": round(pl["kw"], 2), "izmereno_kw_godisno": round(pl["kw_god"], 2), "meseci": pl["meseci"],
                     "stari_v4": round((pl["kw"] - led0) / W_EFF), "stari_v5": round(est[pl["mesto"]]),
                     "stari_najmalku": None if nmin is None else round(nmin),
                     "zabeleska": "само стари светилки" if not pl["led"] else
                     "LED трошат повеќе од измереното – послаби или помалку LED" if pl["kw"] < led0 else
                     "најмалиот можен број (LED 64,5 W не се доволни)" if nmin and round(est[pl["mesto"]]) <= round(nmin) else ""})
    groups = []
    for g in A.GROUPS + ["Вкупно"]:
        rs = [r for r in rows if g == "Вкупно" or r["opstina"] == g]
        led = sum(r["led"] for r in rs)
        n = sum(r["stari_v5"] for r in rs)
        groups.append({"opstina": g, "led": led, "led_kw_popis": round(sum(r["led_kw_popis"] for r in rs), 1), "led_kw_v5": round(sum(r["led_kw_v5"] for r in rs), 1),
                       "udel_zameni": None if g == "Вкупно" else round(p[g], 3), "procenka_udel": None if g == "Вкупно" else share[g],
                       "procenka_stari": round(sum(targets.values()) if g == "Вкупно" else targets[g]), "stari_v4": sum(r["stari_v4"] for r in rs),
                       "stari_v5": n, "udel_v5": round(n / (n + led), 3), "samo_stari_sela": sum(r["stari_v5"] for r in rs if not r["led"]),
                       "zabeleska": "" if g == "Вкупно" else flags[g]})
    checks = {"led_vo_analiza": sum(r["led"] for r in rows), "q_grad": q, "p_staroec": p_st, "referenca": REFERENCE}
    return months, rows, groups, scenarios, control(places, p, q), checks


def write(rows, groups, scenarios, ctrl, checks):
    for name, data in (("v5_po_mesto.csv", rows), ("v5_po_opstina.csv", groups), ("v5_scenarija.csv", scenarios), ("v5_kontrolni_sela.csv", ctrl)):
        with open(A.IZLEZ / name, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=list(data[0].keys()))
            w.writeheader(); w.writerows(data)

    bold, head = Font(bold=True), PatternFill("solid", fgColor="E8EEF2")
    wb = Workbook()

    def sheet(title, header, data, widths):
        ws = wb.create_sheet(title)
        ws.append(header)
        for c in ws[1]:
            c.font, c.fill, c.alignment = bold, head, Alignment(wrap_text=True, vertical="top")
        for r in data:
            ws.append(r)
        for i, wd in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = wd
        ws.freeze_panes = "C2"
        for row in ws.iter_rows(min_row=2):
            for c in row:
                if isinstance(c.value, float):
                    c.number_format = "0.0%" if c.column_letter in ("E", "F", "J") and title == "По општина" else "#,##0.00"
                elif isinstance(c.value, int):
                    c.number_format = "#,##0"
        return ws

    tot = groups[-1]
    q = checks["q_grad"]
    city_strong = 455
    ws = wb.active
    ws.title = "Пристап"
    lines = [
        ("Анализа на стари светилки – верзија 5 (LED моќност според табелата на моќности)", True),
        ("Цел: бројот на стари (несменети) светилки. Бројот 1.500 е само репер.", False),
        ("", False),
        ("Зошто нова верзија", True),
        ("Во Староец (29 LED + 26 стари, избројани) потрошувачката е двојно поголема од тоа што би го трошеле светилките според пописот.", False),
        ("Се објаснува ако околу 21 од 29 LED се 64,5 W наместо 25,9 W: LED се вградувани „што имало“, а не според пописот.", False),
        ("", False),
        ("Претпоставки", True),
        ("Бројот на LED по место е како во пописот (6.500); моќноста се менува со замени во табелата 12,1 / 20,9 / 25,9 / 32,1 / 64,5 W.", False),
        ("Села: дел p од LED се 64,5 W (p < 0: дел се за еден степен послаби). Град: дел q од светилките од 64,5 W се 32,1 W.", False),
        ("Стари светилки 10 W, 80 % активни; загуби 10 %; астрономски часовник; украсување 15.12–15.01 на 12 броила во градот.", False),
        ("Измерена моќност по место = медијана на месечните kWh ÷ ноќни часови (неправилните читања не влијаат).", False),
        ("Удел p по поранешна општина: таков што бројот на стари да одговара на вашата проценка (Зајас 10 %, Осломеј 10 %, Кичево 20 %, Другово 40 %, Вранештица 70 %).", False),
        (f"Град: q = {q:.0%} (околу {round(q * city_strong)} од {city_strong} светилки од 64,5 W се 32,1 W), за градот да има {CITY_TARGET} стари.", False),
        ("Селата без LED (Раштани, Душегубица, Орланци ...) не зависат од LED: бројот се добива директно.", False),
        ("", False),
        ("Резултат", True),
        (f"Вкупно стари светилки: {tot['stari_v5']:,} (репер 1.500; v4 со пописот: {tot['stari_v4']:,}).".replace(",", "."), False),
        (f"Вградена LED моќност: {tot['led_kw_v5']:,} kW наместо {tot['led_kw_popis']:,} kW според пописот.".replace(",", "."), False),
        ("", False),
        ("Ниво на точност", True),
        ("Сметките ја покажуваат само вкупната моќност по мерно место: LED и стари не можат да се разделат само од нив.", False),
        ("Затоа вкупниот број по поранешна општина го следи вашата проценка; анализата ја дава распределбата по место, најмалиот можен број и проверки.", False),
        ("Проверки: контролните села (лист „Контролни села“), независна пресметка со удел од Староец и сценарија (лист „Сценарија“).", False),
    ]
    for text, is_head in lines:
        ws.append([text])
        if is_head:
            ws[ws.max_row][0].font = Font(bold=True, size=13 if ws.max_row == 1 else 11)
    ws.column_dimensions["A"].width = 140

    sheet("Контролни села", ["Место", "LED избројани", "Стари избројани", "LED попис", "Измерено kW", "LED kW (попис)", "Потребно LED kW", "LED од 64,5 W",
                             "Заклучок", "v4: стари", "v5: стари"],
          [list(r.values()) for r in ctrl], [12, 9, 9, 9, 10, 10, 10, 9, 60, 9, 9])
    sheet("По општина", ["Поранешна општина", "LED", "LED kW (попис)", "LED kW (v5)", "Удел замени p", "Ваша проценка (удел)", "Ваша проценка (стари)", "v4: стари",
                         "v5: стари", "v5: удел стари", "од тоа села без LED", "Забелешка"],
          [list(g.values()) for g in groups], [18, 8, 10, 10, 10, 10, 10, 9, 9, 9, 10, 28])
    sheet("По место", ["Поранешна општина", "Населено место", "LED", "LED kW (попис)", "LED kW (v5)", "Измерено kW (медијана)", "Измерено kW (годишно)", "Месеци",
                       "v4: стари", "v5: стари", "Најмалку стари (сите LED 64,5 W)", "Забелешка"],
          [list(r.values()) for r in rows], [12, 22, 7, 9, 9, 10, 10, 7, 9, 9, 11, 50])
    sheet("Сценарија", ["Сценарио"] + A.GROUPS + ["Вкупно", "од тоа градот"],
          [[s["scenario"]] + [s[g] for g in A.GROUPS + ["Вкупно"]] + [s["grad"]] for s in scenarios], [62, 9, 9, 9, 11, 9, 9, 11])
    wb.save(A.IZLEZ / "Анализа_светилки_v5.xlsx")


def main():
    months, rows, groups, scenarios, ctrl, checks = analyse()
    write(rows, groups, scenarios, ctrl, checks)
    for r in ctrl:
        print(r)
    for g in groups:
        print(g)
    for s in scenarios:
        print(f"{s['scenario'][:70]:70} " + " ".join(f"{s[g]:6}" for g in A.GROUPS + ["Вкупно"]) + f"  град {s['grad']}")
    print(checks)


if __name__ == "__main__":
    main()
