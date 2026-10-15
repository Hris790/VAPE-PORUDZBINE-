"""Izveštaj prodaje (dashboard za direktore).

Isti račun i isti izgled kao lokalna skripta „Dashboard_Akcija_Vape", samo bez
pitanja u konzoli — sve što se ranije kucalo (potpun/nepotpun, poslednji mesec u
profitabilnosti, promene cene u toku meseca) sada dolazi kao parametar iz aplikacije.

Glavna funkcija:
    generisi_izvestaj_prodaje(sis_file, tro_file, potpun=True,
                              iskljuci_poslednji=False, promene=None)
    -> (html, xlsx_bytes, mesec_label, prodaja_summary)

`promene` je lista rečnika:
    {"sistem": str, "grupa": str, "godina": int, "mesec": int,
     "dan": int, "stara": float, "nova": float}
"""

# Oznaka verzije modula — vidi se u aplikaciji (Objava izveštaja → Izveštaj prodaje),
# da se odmah zna da li je na GitHub-u najnoviji fajl.
VERZIJA = "3.1 · 15.10.2026 (Excel kao tabela na ekranu, boje cena, legenda iznad tabele)"

import calendar
import html as html_mod
import io
import json
import re
from collections import Counter

import pandas as pd

# =====================================================================
# BOJE PO CENOVNIM NIVOIMA
# =====================================================================
RAMPA_ANCHORS = ['#90EE90', '#FFD1DC', '#FFB6C1', '#FF69B4', '#C71585', '#DC2626']
TOLERANCIJA_CENE = 25          # RSD — cene unutar ovog raspona = isti nivo

GRUPE_BEZ_BOJE = {'NERD 1000', 'GOAT', 'KILLU', 'KILLA', 'ZVAKE'}
_GRUPE_BEZ_BOJE_N = {' '.join(g.strip().upper().split()) for g in GRUPE_BEZ_BOJE}

REDOVNE_CENE = {'NERD 2000': 1390, 'HQD 1000': 890, 'SYX': 490}
_REDOVNE_N = {k.strip().upper(): v for k, v in REDOVNE_CENE.items()}

MAPA_MESECI = {1: 'Jan', 2: 'Feb', 3: 'Mar', 4: 'Apr', 5: 'Maj', 6: 'Jun', 7: 'Jul',
               8: 'Avg', 9: 'Sep', 10: 'Okt', 11: 'Nov', 12: 'Dec'}
MAPA_MESECI_PUN = {1: 'Januar', 2: 'Februar', 3: 'Mart', 4: 'April', 5: 'Maj', 6: 'Jun',
                   7: 'Jul', 8: 'Avgust', 9: 'Septembar', 10: 'Oktobar', 11: 'Novembar',
                   12: 'Decembar'}

TROSAK_NAZIVI = {
    'trosak transporta': 'Trošak transporta', 'trosak advokata': 'Trošak advokata',
    'tehnomedia/gigatron': 'Tehnomedia / Gigatron', 'trosak vozila': 'Trošak vozila',
    'reprezentacija': 'Reprezentacija', 'trosak programa': 'Trošak programa',
    'racuni + banka+ osiguranje': 'Računi + banka + osiguranje', 'plate': 'Plate',
    'trosak laboratorije': 'Trošak laboratorije', 'promotivni troskovi': 'Promotivni troškovi',
    'troškovi kanc/ materijala': 'Troškovi kanc. materijala', 'ostali troškovi': 'Ostali troškovi',
}

POSMATRANE_GRUPE = ['HQD 1000', 'NERD 2000']
TRAZENE_GRUPE = ['HQD 1000', 'NERD 2000', 'SYX']


def bez_boje(gnm):
    n = ' ' + ' '.join(str(gnm).strip().upper().replace('-', ' ').split()) + ' '
    return any((' ' + t + ' ') in n for t in _GRUPE_BEZ_BOJE_N)


def _hex2rgb(h):
    h = h.lstrip('#')
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def _rgb2hex(t):
    return '#%02X%02X%02X' % tuple(max(0, min(255, int(round(v)))) for v in t)


def _mix(c1, c2, f):
    a, b = _hex2rgb(c1), _hex2rgb(c2)
    return _rgb2hex([a[i] + (b[i] - a[i]) * f for i in range(3)])


def napravi_rampu(n):
    """n boja: od zelene (najviša cena) do crvene (najniža)."""
    if n <= 0:
        return []
    if n == 1:
        return [RAMPA_ANCHORS[0]]
    if n == 2:
        return [RAMPA_ANCHORS[0], RAMPA_ANCHORS[4]]
    out = []
    for i in range(n):
        t = i / (n - 1) * (len(RAMPA_ANCHORS) - 1)
        lo = int(t)
        hi = min(lo + 1, len(RAMPA_ANCHORS) - 1)
        out.append(_mix(RAMPA_ANCHORS[lo], RAMPA_ANCHORS[hi], t - lo))
    return out


def klasteruj_cene(cene, tol=TOLERANCIJA_CENE):
    ce = sorted(float(c) for c in cene if c and c > 0)
    if not ce:
        return []
    klasteri = [[ce[0]]]
    for c in ce[1:]:
        if c - klasteri[-1][-1] <= tol:
            klasteri[-1].append(c)
        else:
            klasteri.append([c])
    return [Counter(int(round(v)) for v in k).most_common(1)[0][0] for k in klasteri]


def is_tamna(hexc):
    r, g, b = _hex2rgb(hexc)
    return (0.299 * r + 0.587 * g + 0.114 * b) < 150


def text_color(hexc):
    return '#fff' if is_tamna(hexc) else '#1a1a2e'


def esc(t):
    return html_mod.escape(str(t))


def fmtnum(v):
    return f"{round(v):,}"


def kratak_ukus(naziv):
    n = str(naziv).strip()
    n = re.sub(r'(?i)\b(jednokratna|e-?cigareta|nerd|ure[dđ]aj|puffs?|nikotinske|vre[cć]ice|'
               r'bez duvana|sa zasla[dđ]iva[cč]em|hqd|syx|goat|killa|killu|zvake|žvake)\b', ' ', n)
    n = re.sub(r'\(?\s*\d+[.,]?\d*\s*(mg|%)[^)]*\)?', ' ', n)
    n = re.sub(r'\b\d{3,4}\b', ' ', n)
    n = re.sub(r'[()/]', ' ', n)
    n = ' '.join(n.split())
    return n if n else str(naziv).strip()


def pct(curr, ref):
    if not ref:
        return 0.0
    return round((curr - ref) / ref * 100, 1)


def fmt_diff(v):
    return ('+' if v >= 0 else '') + fmtnum(v)


def fmt_pct_signed(v):
    if v is None:
        return "—"
    return ('+' if v >= 0 else '') + f"{v}%"


def color_pct(v):
    if v is None:
        return '#8b90a5'
    return '#16a34a' if v > 0 else ('#dc2626' if v < 0 else '#8b90a5')


def badge(v):
    if v is None:
        return ''
    color = '#16a34a' if v >= 0 else '#dc2626'
    bg = 'rgba(22,163,74,0.08)' if v >= 0 else 'rgba(220,38,38,0.08)'
    arr = '▲' if v >= 0 else '▼'
    rec = 'rast od' if v >= 0 else 'pad od'
    return (f'<span style="background:{bg};color:{color};padding:2px 8px;border-radius:6px;'
            f'font-weight:700;font-family:\'IBM Plex Mono\',monospace;font-size:13px">'
            f'{arr} {rec} {abs(v)}%</span>')


def _ucitaj_sheet(xls, kljucne, obavezno=True):
    """Nađi list po delu naziva (bez obzira na razmake i velika slova)."""
    sm = {str(s).strip().lower(): s for s in xls.sheet_names}
    for kw in kljucne:
        for nl, no in sm.items():
            if kw in nl:
                return pd.read_excel(xls, sheet_name=no)
    if obavezno:
        raise ValueError("U fajlu nema lista: " + " / ".join(kljucne))
    return pd.DataFrame()


# =====================================================================
# GLAVNA FUNKCIJA
# =====================================================================
def generisi_izvestaj_prodaje(sis_file, tro_file, potpun=True, iskljuci_poslednji=False,
                              promene=None):
    xls = pd.ExcelFile(sis_file)
    df = _ucitaj_sheet(xls, ['tabela'])
    df.columns = df.columns.astype(str).str.strip()

    df_mag = _ucitaj_sheet(xls, ['zalihe magacin', 'magacin'], obavezno=False)
    if len(df_mag):
        df_mag.columns = df_mag.columns.astype(str).str.strip()
        for c in df_mag.columns:
            if 'kol' in c.lower().strip():
                df_mag.rename(columns={c: 'KOL'}, inplace=True)
                break

    df_troskovi = _ucitaj_sheet(pd.ExcelFile(tro_file), ['tabela'])
    df_troskovi.columns = df_troskovi.columns.str.strip()
    df_troskovi['SISTEM'] = df_troskovi['SISTEM'].astype(str).str.strip()
    df_troskovi['Mesec'] = pd.to_numeric(df_troskovi['Mesec'], errors='coerce').fillna(0).astype(int)
    df_troskovi['Godina'] = pd.to_numeric(df_troskovi['Godina'], errors='coerce').fillna(0).astype(int)

    trosak_kolone = [c for c in df_troskovi.columns if c not in ['SISTEM', 'Mesec', 'Godina']]
    mapa_troskova = {}
    for _, row in df_troskovi.iterrows():
        key = f"{row['SISTEM']}|{int(row['Godina'])}-{int(row['Mesec'])}"
        mapa_troskova[key] = {c: (float(row[c]) if pd.notna(row[c]) else 0) for c in trosak_kolone}

    col_nacin = next((c for c in df.columns
                      if "PLACANJA" in c.upper().replace('Č', 'C').replace('Ć', 'C')), None)
    mapa_placanja = {}
    if col_nacin:
        temp = df[['SISTEM', col_nacin]].dropna()
        mapa_placanja = dict(zip(temp['SISTEM'], temp[col_nacin]))

    df['Mesec'] = pd.to_numeric(df['Mesec'], errors='coerce').fillna(0).astype(int)
    df['Godina'] = pd.to_numeric(df['Godina'], errors='coerce').fillna(0).astype(int)
    df_clean = df[(df['Godina'] >= 2025) & (df['Mesec'] >= 1)].copy()

    col_dod = next((c for c in df_clean.columns
                    if c.strip().upper().startswith('DODATNI MESECNI TROSAK')), None)
    mapa_dodatni = {}
    if col_dod:
        dod_grouped = df_clean.groupby(['SISTEM', 'Godina', 'Mesec'])[col_dod].sum().reset_index()
        for _, row in dod_grouped.iterrows():
            v = float(row[col_dod]) if pd.notna(row[col_dod]) else 0
            if v > 0:
                mapa_dodatni[f"{row['SISTEM']}|{int(row['Godina'])}-{int(row['Mesec'])}"] = v

    periodi = df_clean[['Godina', 'Mesec']].drop_duplicates().sort_values(['Godina', 'Mesec']).values.tolist()
    if not periodi:
        raise ValueError("U tabeli nema nijednog meseca (Godina/Mesec).")
    nazivi = [f"{MAPA_MESECI[int(m)]} {str(int(g))[-2:]}" for g, m in periodi]
    poslednji_g, poslednji_m = periodi[-1]
    poslednji_naziv = f"{MAPA_MESECI[int(poslednji_m)]} {int(poslednji_g)}"

    if potpun and iskljuci_poslednji and len(periodi) > 1:
        periodi_profit, nazivi_profit = periodi[:-1], nazivi[:-1]
        badge_html = f'<span class="badge br">{poslednji_naziv}: ISKLJUČEN iz profitabilnosti</span>'
    else:
        periodi_profit, nazivi_profit = periodi[:], nazivi[:]
        badge_html = (f'<span class="badge bg">{poslednji_naziv}: UKLJUČEN</span>'
                      if potpun else '')
    periodi_drv, nazivi_drv = periodi[:], nazivi[:]
    num_s, num_p, num_d = len(nazivi), len(nazivi_profit), len(nazivi_drv)

    sistemi_lista = sorted(df_clean['SISTEM'].dropna().unique())
    sve_grupe = sorted(df_clean['Grupa artikla'].dropna().astype(str).str.strip().unique())

    # ---------- promene cene u toku meseca ----------
    promene_cena, promene_lista = {}, []
    for p in (promene or []):
        try:
            s = str(p.get("sistem", "")).strip()
            gr = str(p.get("grupa", "")).strip()
            gg = int(p.get("godina"))
            mm = int(p.get("mesec"))
            dan = int(p.get("dan"))
            stara = float(p.get("stara"))
            nova = float(p.get("nova"))
        except Exception:
            continue
        if not s or not gr or stara <= 0 or nova <= 0:
            continue
        dmax = calendar.monthrange(gg, mm)[1]
        dan = max(1, min(dmax, dan))
        promene_cena[f"{s}|{gr}|{gg}-{mm}"] = {'dan': dan, 'stara': stara, 'nova': nova}
        promene_lista.append({
            'sistem': s, 'grupa': gr, 'mesec_naziv': f"{MAPA_MESECI[mm]} {str(gg)[-2:]}",
            'dan': dan, 'mesec': mm, 'godina': gg, 'stara': stara, 'nova': nova,
            'smer': 'dole' if nova < stara else ('gore' if nova > stara else 'isto')})

    # ---------- palete boja po grupi ----------
    _cene_grupa = {}
    for _gnm_raw, _v in zip(df_clean['Grupa artikla'], df_clean['FINALNA MP']):
        gnm = str(_gnm_raw).strip()
        if gnm in ('', 'nan') or bez_boje(gnm):
            continue
        if pd.notna(_v) and float(_v) > 0:
            _cene_grupa.setdefault(gnm, []).append(float(_v))
    for _k, _v in promene_cena.items():
        gnm = _k.split('|')[1]
        if bez_boje(gnm):
            continue
        _cene_grupa.setdefault(gnm, []).extend([_v['stara'], _v['nova']])

    grupa_palete, grupa_redovna = {}, {}
    for gnm, lst in _cene_grupa.items():
        nivoi = klasteruj_cene(lst)
        red = _REDOVNE_N.get(gnm.strip().upper())
        if red:
            nivoi = [c for c in nivoi if c <= red + TOLERANCIJA_CENE]
            if not any(abs(c - red) <= TOLERANCIJA_CENE for c in nivoi):
                nivoi.append(red)
            else:
                nivoi = [red if abs(c - red) <= TOLERANCIJA_CENE else c for c in nivoi]
            grupa_redovna[gnm] = red
        nivoi = sorted(set(nivoi), reverse=True)
        grupa_palete[gnm] = list(zip(nivoi, napravi_rampu(len(nivoi))))

    def get_color(price, grupa=None):
        if not price or price <= 0 or grupa is None:
            return None
        if bez_boje(grupa):
            return None
        nivoi = grupa_palete.get(str(grupa).strip())
        if not nivoi:
            return None
        best = min(nivoi, key=lambda x: abs(x[0] - price))
        return best[1] if abs(best[0] - price) <= TOLERANCIJA_CENE + 10 else None

    def cell_c(bg, val, i):
        cls = f"n mcell m{i}"
        if not bg:
            return f'<td class="{cls}">{fmtnum(val)}</td>'
        return (f'<td class="{cls}" style="background:{bg};color:{text_color(bg)};'
                f'font-weight:600">{fmtnum(val)}</td>')

    def cell_split(val, i, grupa, pcn):
        dan = int(pcn['dan'])
        stara, nova = pcn['stara'], pcn['nova']
        g, m = periodi[i]
        dana_u_mesecu = calendar.monthrange(int(g), int(m))[1]
        p = (dan - 1) / dana_u_mesecu * 100
        p = max(6, min(94, round(p)))
        c1 = get_color(stara, grupa) or '#e8ecf4'
        c2 = get_color(nova, grupa) or '#cbd5e1'
        dominantna = c1 if p >= 50 else c2
        bg = (f"linear-gradient(90deg,{c1} 0%,{c1} {p}%,"
              f"#1a1a2e {p}%,#1a1a2e calc({p}% + 2px),"
              f"{c2} calc({p}% + 2px),{c2} 100%)")
        tip = (f"Cena menjana {dan}.{int(m)}.{int(g)}. — do tada {fmtnum(stara)} RSD, "
               f"od tada {fmtnum(nova)} RSD")
        return (f'<td class="n mcell m{i} pc-cell" title="{esc(tip)}" '
                f'style="background:{bg};color:{text_color(dominantna)};'
                f'font-weight:700">{fmtnum(val)}</td>')

    def mix_nivoi(g_data, g, m):
        mask = (g_data['Godina'] == int(g)) & (g_data['Mesec'] == int(m))
        sub = g_data.loc[mask]
        agg = {}
        for _, rr in sub.iterrows():
            c = rr.get('FINALNA MP')
            if pd.isna(c) or float(c) <= 0:
                continue
            c = int(round(float(c)))
            q = rr.get('Prodata kolicina ka krajnjem kupcu')
            q = float(q) if pd.notna(q) else 0.0
            a = str(rr.get('Artikl', '')).strip()
            d = agg.setdefault(c, {'qty': 0.0, 'n': 0, 'arts': []})
            d['qty'] += q
            d['n'] += 1
            if a and a.lower() != 'nan':
                d['arts'].append(kratak_ukus(a))
        return [(c, d['qty'], d['n'], d['arts'])
                for c, d in sorted(agg.items(), key=lambda kv: -kv[0])]

    def cell_mix(val, i, grupa, nivoi):
        uk_qty = sum(x[1] for x in nivoi)
        udeli = [x[1] for x in nivoi] if uk_qty > 0 else [float(x[2]) for x in nivoi]
        uk = sum(udeli) or 1.0
        sirine = [max(10.0, u / uk * 100.0) for u in udeli]
        norm = sum(sirine)
        sirine = [w / norm * 100.0 for w in sirine]
        delovi, poc, poz, naj, najw = [], "0%", 0.0, None, -1
        for idx, ((cena, qty, na, arts), w) in enumerate(zip(nivoi, sirine)):
            col = get_color(cena, grupa) or '#e8ecf4'
            if w > najw:
                najw, naj = w, col
            poz += w
            kraj = "100%" if idx == len(sirine) - 1 else f"{poz:.2f}%"
            delovi.append(f"{col} {poc},{col} {kraj}")
            if idx < len(sirine) - 1:
                delovi.append(f"#1a1a2e {kraj},#1a1a2e calc({kraj} + 2px)")
                poc = f"calc({kraj} + 2px)"
        bg = "linear-gradient(90deg," + ",".join(delovi) + ")"
        tip_red = []
        for (cena, qty, na, arts) in nivoi:
            imena = ", ".join(arts[:5]) + ("…" if len(arts) > 5 else "")
            tip_red.append(f"{fmtnum(cena)} RSD — {na} ukusa, {fmtnum(qty)} kom"
                           + (f": {imena}" if imena else ""))
        tip = "Različite cene u grupi:\n" + "\n".join(tip_red)
        return (f'<td class="n mcell m{i} pc-cell mix-cell" title="{esc(tip)}" '
                f'style="background:{bg};color:{text_color(naj or "#e8ecf4")};'
                f'font-weight:700">{fmtnum(val)}</td>')

    # ---------- podaci za JS ----------
    prodaja_data_js = {}
    for sistem in sistemi_lista:
        s_data = df_clean[df_clean['SISTEM'] == sistem]
        prodaja_data_js[sistem] = {}
        for grupa in sorted(s_data['Grupa artikla'].dropna().astype(str).str.strip().unique()):
            g_data = s_data[s_data['Grupa artikla'].astype(str).str.strip() == grupa]
            gvals = []
            for g, m in periodi:
                mask = (g_data['Godina'] == int(g)) & (g_data['Mesec'] == int(m))
                gvals.append(round(float(g_data.loc[mask, 'Prodata kolicina ka krajnjem kupcu'].sum())))
            prodaja_data_js[sistem][grupa] = gvals

    trosak_names_list = ['Troškovi marketinga']
    trosak_ids_list = ['mkt']
    for ki, kat in enumerate(trosak_kolone):
        trosak_ids_list.append(f"t{ki}")
        trosak_names_list.append(TROSAK_NAZIVI.get(kat, kat))
    trosak_ids_list.append('dod')
    trosak_names_list.append('Dodatni mesečni trošak')

    profit_data_js = {}
    for sistem in sistemi_lista:
        s_data = df_clean[df_clean['SISTEM'] == sistem]
        is_f = str(mapa_placanja.get(sistem, None)) in ['1', '1.0']
        pn, mn, dn = [], [], []
        for g, m in periodi_profit:
            mask = (s_data['Godina'] == int(g)) & (s_data['Mesec'] == int(m))
            pn.append(round(float(s_data.loc[mask, 'PROFIT3' if is_f else 'Profit'].sum())))
            mn.append(round(float(s_data.loc[mask, 'MESECNI TROSAK1'].sum())))
            dn.append(round(mapa_dodatni.get(f"{sistem}|{int(g)}-{int(m)}", 0)))
        profit_data_js[sistem] = {'profit': pn, 'mkt': mn, 'dod': dn}
        for ki, kat in enumerate(trosak_kolone):
            profit_data_js[sistem][f"t{ki}"] = [
                round(mapa_troskova.get(f"{sistem}|{int(g)}-{int(m)}", {}).get(kat, 0))
                for g, m in periodi_profit]

    drv_data_js = {}
    for sistem in sistemi_lista:
        s_data = df_clean[df_clean['SISTEM'] == sistem]
        pn, mn, dn = [], [], []
        for g, m in periodi_drv:
            mask = (s_data['Godina'] == int(g)) & (s_data['Mesec'] == int(m))
            pn.append(round(float(s_data.loc[mask, 'Profit'].sum())))
            mn.append(round(float(s_data.loc[mask, 'MESECNI TROSAK1'].sum())))
            dn.append(round(mapa_dodatni.get(f"{sistem}|{int(g)}-{int(m)}", 0)))
        drv_data_js[sistem] = {'profit': pn, 'mkt': mn, 'dod': dn}
        for ki, kat in enumerate(trosak_kolone):
            drv_data_js[sistem][f"t{ki}"] = [
                round(mapa_troskova.get(f"{sistem}|{int(g)}-{int(m)}", {}).get(kat, 0))
                for g, m in periodi_drv]

    # ---------- magacin ----------
    last3 = periodi[-3:]
    df_art = df_clean[df_clean['Artikl'].notna()].copy()
    df_art['Artikl'] = df_art['Artikl'].astype(str).str.strip()
    art_monthly_sales = {}
    for art in df_art['Artikl'].unique():
        a_data = df_art[df_art['Artikl'] == art]
        sales = []
        for g, m in last3:
            mask = (a_data['Godina'] == int(g)) & (a_data['Mesec'] == int(m))
            sales.append(float(a_data.loc[mask, 'Prodata kolicina ka krajnjem kupcu'].sum()))
        art_monthly_sales[art] = round(sum(sales) / len(sales)) if sales else 0

    def get_grupa(art_name):
        if 'HQD' in art_name.upper():
            return 'HQD 1000'
        if '2000' in art_name:
            return 'NERD 2000'
        if 'E-cigareta' in art_name or '1000' in art_name:
            return 'NERD 1000'
        return 'Ostalo'

    mag_data, total_mag, total_avg = [], 0, 0
    for _, row in df_mag.iterrows():
        art_name = str(row.get('Naziv artikla', '')).strip()
        try:
            kol = int(row['KOL'])
        except Exception:
            continue
        avg_sale = art_monthly_sales.get(art_name, 0)
        daily = avg_sale / 30 if avg_sale > 0 else 0
        days = round(kol / daily) if daily > 0 else 9999
        mag_data.append((art_name, kol, avg_sale, days))
        total_mag += kol
        total_avg += avg_sale
    mag_data.sort(key=lambda x: (get_grupa(x[0]), x[0]))

    mag_rows = []
    for art_name, kol, avg_sale, days in mag_data:
        grupa = get_grupa(art_name)
        if days <= 30:
            dcls = 'style="background:#fee2e2;color:#b91c1c;font-weight:700"'
        elif days <= 60:
            dcls = 'style="background:#fef3c7;color:#92400e;font-weight:600"'
        elif days <= 90:
            dcls = 'style="background:#e0f2fe;color:#075985;font-weight:600"'
        else:
            dcls = 'style="color:var(--grn);font-weight:600"'
        days_str = f"{days}" if days < 9999 else "∞"
        months_str = f"{days / 30:.1f}" if days < 9999 else "∞"
        mag_rows.append(
            f'<tr class="mag-row" data-grupa="{esc(grupa)}">'
            f'<td style="font-size:10px;color:var(--t2)">{esc(art_name)}</td>'
            f'<td class="nb" style="color:var(--t2);font-size:10px">{esc(grupa)}</td>'
            f'<td class="nb">{fmtnum(kol)}</td><td class="nb">{fmtnum(avg_sale)}</td>'
            f'<td class="nb" {dcls}>{days_str}</td><td class="nb" {dcls}>{months_str}</td></tr>')
    total_daily = total_avg / 30 if total_avg > 0 else 0
    total_days = round(total_mag / total_daily) if total_daily > 0 else 9999
    total_months_str = f"{total_days / 30:.1f}" if total_days < 9999 else "∞"
    if mag_data:
        mag_rows.append(
            f'<tr class="totalrow"><td class="total-label">UKUPNO MAGACIN</td><td></td>'
            f'<td class="nb total-cell">{fmtnum(total_mag)}</td>'
            f'<td class="nb total-cell">{fmtnum(total_avg)}</td>'
            f'<td class="nb total-cell" style="font-weight:800">{total_days}</td>'
            f'<td class="nb total-cell" style="font-weight:800">{total_months_str}</td></tr>')

    mag_grupe_summary = {}
    for art_name, kol, avg_sale, days in mag_data:
        g = get_grupa(art_name)
        d = mag_grupe_summary.setdefault(g, {'kol': 0, 'avg': 0})
        d['kol'] += kol
        d['avg'] += avg_sale

    # ---------- zalihe po sistemima ----------
    def get_last_zalihe(data, periodi_list):
        for g, m in reversed(periodi_list):
            mask = (data['Godina'] == int(g)) & (data['Mesec'] == int(m))
            zal = data.loc[mask, 'Zalihe'].sum() if 'Zalihe' in data.columns else 0
            if pd.notna(zal) and zal > 0:
                return round(float(zal)), f"{MAPA_MESECI[int(m)]} {str(int(g))[-2:]}"
        return 0, ""

    zal_rows, si, zal_grand = [], 0, 0
    for sistem in sistemi_lista:
        si += 1
        sid = f"zs{si}"
        s_data = df_clean[df_clean['SISTEM'] == sistem]
        s_zal, s_per = get_last_zalihe(s_data, periodi)
        zal_grand += s_zal
        zal_rows.append(
            f'<tr class="sr" data-sistem="{esc(sistem)}" onclick="tog(\'{sid}\')" style="cursor:pointer">'
            f'<td><button class="be" id="b-{sid}">+</button><span class="sn">{esc(sistem)}</span></td>'
            f'<td class="nb">{fmtnum(s_zal)}</td>'
            f'<td class="n" style="color:var(--t3);font-size:10px">{s_per}</td></tr>')
        gi = 0
        for grupa in sorted(s_data['Grupa artikla'].dropna().astype(str).str.strip().unique()):
            gi += 1
            gid = f"{sid}g{gi}"
            g_data = s_data[s_data['Grupa artikla'].astype(str).str.strip() == grupa]
            g_zal, g_per = get_last_zalihe(g_data, periodi)
            zal_rows.append(
                f'<tr class="gr hidden" data-p="{sid}" data-sistem="{esc(sistem)}" '
                f'data-grupa="{esc(grupa)}" onclick="tog(\'{gid}\');event.stopPropagation()" '
                f'style="cursor:pointer"><td style="padding-left:28px">'
                f'<button class="be beg" id="b-{gid}">+</button><span class="gn">{esc(grupa)}</span></td>'
                f'<td class="nb">{fmtnum(g_zal)}</td>'
                f'<td class="n" style="color:var(--t3);font-size:10px">{g_per}</td></tr>')
            for art in sorted(g_data['Artikl'].dropna().astype(str).str.strip().unique()):
                if not art:
                    continue
                a_data = g_data[g_data['Artikl'].astype(str).str.strip() == art]
                a_zal, a_per = get_last_zalihe(a_data, periodi)
                zal_rows.append(
                    f'<tr class="ar hidden" data-p="{gid}" data-sistem="{esc(sistem)}" '
                    f'data-grupa="{esc(grupa)}"><td class="an">{esc(art)}</td>'
                    f'<td class="n" style="color:#7c8494;font-size:10px">'
                    f'{fmtnum(a_zal) if a_zal > 0 else ""}</td>'
                    f'<td class="n" style="color:var(--t3);font-size:9px">{a_per}</td></tr>')
        zal_rows.append(f'<tr class="sep" data-sistem="{esc(sistem)}"><td colspan="999"></td></tr>')
    zal_rows.append(
        f'<tr class="totalrow" id="zalihe-total"><td class="total-label">TOTAL</td>'
        f'<td class="nb total-cell" style="font-size:13px">{fmtnum(zal_grand)}</td><td></td></tr>')

    # ---------- tabela PRODAJA ----------
    prodaja_rows, grand_totals, mix_lista, si = [], [0] * num_s, [], 0
    xls_prodaja = []          # ista struktura kao tabela na ekranu — za Excel
    for sistem in sistemi_lista:
        si += 1
        sid = f"ps{si}"
        s_data = df_clean[df_clean['SISTEM'] == sistem]
        vals = []
        for g, m in periodi:
            mask = (s_data['Godina'] == int(g)) & (s_data['Mesec'] == int(m))
            vals.append(round(float(s_data.loc[mask, 'Prodata kolicina ka krajnjem kupcu'].sum())))
        for i in range(num_s):
            grand_totals[i] += vals[i]
        r = (f'<tr class="sr" data-sistem="{esc(sistem)}" onclick="tog(\'{sid}\')" style="cursor:pointer">'
             f'<td><button class="be" id="b-{sid}">+</button><span class="sn">{esc(sistem)}</span></td>')
        for i, v in enumerate(vals):
            r += f'<td class="nb mcell m{i}">{fmtnum(v)}</td>'
        r += f'<td class="nt totcell">{fmtnum(sum(vals))}</td></tr>'
        prodaja_rows.append(r)
        _xs = {"sistem": sistem, "vals": list(vals), "grupe": []}
        xls_prodaja.append(_xs)
        gi = 0
        for grupa in sorted(s_data['Grupa artikla'].dropna().astype(str).str.strip().unique()):
            gi += 1
            gid = f"{sid}g{gi}"
            g_data = s_data[s_data['Grupa artikla'].astype(str).str.strip() == grupa]
            gvals, gnivoi = [], []
            for g, m in periodi:
                mask = (g_data['Godina'] == int(g)) & (g_data['Mesec'] == int(m))
                gvals.append(round(float(g_data.loc[mask, 'Prodata kolicina ka krajnjem kupcu'].sum())))
                gnivoi.append(mix_nivoi(g_data, g, m))
            ima_pc = any(f"{sistem}|{grupa}|{int(periodi[i][0])}-{int(periodi[i][1])}" in promene_cena
                         for i in range(num_s))
            ima_mix = any(len(nv) > 1 for nv in gnivoi)
            pc_mark = (' <span class="pc-badge" title="U ovoj grupi je cena menjana u toku meseca">⇄</span>'
                       if ima_pc else '')
            if ima_mix:
                pc_mark += (' <span class="mix-badge" title="U ovoj grupi ukusi nisu svi na istoj ceni">'
                            '◧</span>')
            r = (f'<tr class="gr hidden" data-p="{sid}" data-sistem="{esc(sistem)}" '
                 f'data-grupa="{esc(grupa)}" onclick="tog(\'{gid}\');event.stopPropagation()" '
                 f'style="cursor:pointer"><td style="padding-left:28px">'
                 f'<button class="be beg" id="b-{gid}">+</button>'
                 f'<span class="gn">{esc(grupa)}</span>{pc_mark}</td>')
            _xboje, _xnote = [], []
            for i, v in enumerate(gvals):
                key_pc = f"{sistem}|{grupa}|{int(periodi[i][0])}-{int(periodi[i][1])}"
                nv = gnivoi[i]
                if key_pc in promene_cena:
                    r += cell_split(v, i, grupa, promene_cena[key_pc])
                    _pcn = promene_cena[key_pc]
                    _c1 = get_color(_pcn['stara'], grupa)
                    _c2 = get_color(_pcn['nova'], grupa)
                    _xboje.append(_c2 or _c1)
                    _xnote.append("Cena menjana " + str(int(_pcn['dan'])) + "."
                                  + str(int(periodi[i][1])) + "." + str(int(periodi[i][0]))
                                  + ". — do tada " + fmtnum(_pcn['stara']) + " RSD, od tada "
                                  + fmtnum(_pcn['nova']) + " RSD")
                elif len(nv) > 1:
                    r += cell_mix(v, i, grupa, nv)
                    if i == num_s - 1:
                        mix_lista.append({'sistem': sistem, 'grupa': grupa, 'nivoi': nv})
                    _naj = max(nv, key=lambda x: x[1])
                    _xboje.append(get_color(_naj[0], grupa))
                    _xnote.append("Različite cene u grupi: "
                                  + " · ".join(fmtnum(c) + " RSD (" + fmtnum(q) + " kom)"
                                               for c, q, _na, _ar in nv))
                elif len(nv) == 1:
                    r += cell_c(get_color(nv[0][0], grupa), v, i)
                    _xboje.append(get_color(nv[0][0], grupa))
                    _xnote.append(fmtnum(nv[0][0]) + " RSD")
                else:
                    r += cell_c(None, v, i)
                    _xboje.append(None)
                    _xnote.append("")
            r += f'<td class="nb totcell">{fmtnum(sum(gvals))}</td></tr>'
            prodaja_rows.append(r)
            _xg = {"grupa": grupa, "vals": list(gvals), "boje": _xboje,
                   "note": _xnote, "artikli": []}
            _xs["grupe"].append(_xg)
            for art in sorted(g_data['Artikl'].dropna().astype(str).str.strip().unique()):
                if not art:
                    continue
                a_data = g_data[g_data['Artikl'].astype(str).str.strip() == art]
                avals = []
                for g, m in periodi:
                    mask = (a_data['Godina'] == int(g)) & (a_data['Mesec'] == int(m))
                    avals.append(round(float(a_data.loc[mask, 'Prodata kolicina ka krajnjem kupcu'].sum())))
                r = (f'<tr class="ar hidden" data-p="{gid}" data-sistem="{esc(sistem)}" '
                     f'data-grupa="{esc(grupa)}"><td class="an">{esc(art)}</td>')
                for i, v in enumerate(avals):
                    if v > 0:
                        r += (f'<td class="n mcell m{i}" style="color:#7c8494;font-size:10px">'
                              f'{fmtnum(v)}</td>')
                    else:
                        r += f'<td class="n mcell m{i}"></td>'
                r += (f'<td class="n totcell" style="color:#7c8494;font-size:10px">'
                      f'{fmtnum(sum(avals))}</td></tr>')
                prodaja_rows.append(r)
                _xg["artikli"].append({"naziv": art, "vals": list(avals)})
        prodaja_rows.append(f'<tr class="sep" data-sistem="{esc(sistem)}"><td colspan="999"></td></tr>')
    tr_r = '<tr class="totalrow" id="prodaja-total"><td class="total-label">TOTAL</td>'
    for i, v in enumerate(grand_totals):
        tr_r += f'<td class="nb total-cell mcell m{i}">{fmtnum(v)}</td>'
    tr_r += (f'<td class="nt total-cell totcell" style="font-size:13px">'
             f'{fmtnum(sum(grand_totals))}</td></tr>')
    prodaja_rows.append(tr_r)

    # ---------- analitika prodaje ----------
    last_idx = num_s - 1
    curr_total = grand_totals[last_idx]
    prev_total = grand_totals[last_idx - 1] if num_s >= 2 else 0
    n6 = min(6, last_idx)
    avg6_total = round(sum(grand_totals[last_idx - n6:last_idx]) / n6) if n6 > 0 else 0
    yoy_idx = None
    for idx, (g, m) in enumerate(periodi):
        if int(g) == int(poslednji_g) - 1 and int(m) == int(poslednji_m):
            yoy_idx = idx
            break
    yoy_total = grand_totals[yoy_idx] if yoy_idx is not None else None
    pct_mom = pct(curr_total, prev_total)
    pct_6m = pct(curr_total, avg6_total)
    pct_yoy = pct(curr_total, yoy_total) if yoy_total else None
    diff_mom, diff_6m = curr_total - prev_total, curr_total - avg6_total
    diff_yoy = (curr_total - yoy_total) if yoy_total else None
    prev_naziv = nazivi[last_idx - 1] if num_s >= 2 else ""
    yoy_naziv = nazivi[yoy_idx] if yoy_idx is not None else ""

    hp = [f'Prodaja u <b style="color:#2563eb">{poslednji_naziv.lower()}</b> iznosi '
          f'<b style="font-family:\'IBM Plex Mono\',monospace;font-size:15px">'
          f'{fmtnum(curr_total)} kom</b>']
    if num_s >= 2:
        hp.append(f'i beleži {badge(pct_mom)} u odnosu na prethodni mesec '
                  f'<span style="color:#8b90a5">({prev_naziv}: {fmtnum(prev_total)} kom)</span>')
    if n6 > 0:
        hp.append(f', odnosno {badge(pct_6m)} u odnosu na 6-mesečni prosek '
                  f'<span style="color:#8b90a5">({fmtnum(avg6_total)} kom)</span>')
    if pct_yoy is not None:
        hp.append(f'. Poređenjem sa istim mesecom prethodne godine '
                  f'<span style="color:#8b90a5">({yoy_naziv}: {fmtnum(yoy_total)} kom)</span>, '
                  f'rezultat je {badge(pct_yoy)} YoY')
    headline_text = ' '.join(hp) + '.'

    grupa_boje = {
        'HQD 1000': {'main': '#dc2626', 'bg': 'rgba(220,38,38,0.1)'},
        'NERD 2000': {'main': '#2563eb', 'bg': 'rgba(37,99,235,0.1)'},
        'SYX': {'main': '#16a34a', 'bg': 'rgba(22,163,74,0.1)'},
    }
    grupe_data = {}
    for grupa in TRAZENE_GRUPE:
        g_data = df_clean[df_clean['Grupa artikla'].astype(str).str.strip() == grupa]
        vals = []
        for g, m in periodi:
            mask = (g_data['Godina'] == int(g)) & (g_data['Mesec'] == int(m))
            vals.append(round(float(g_data.loc[mask, 'Prodata kolicina ka krajnjem kupcu'].sum())))
        grupe_data[grupa] = vals

    grupe_blocks_html, grupe_chart_data_js = '', {}
    for grupa in TRAZENE_GRUPE:
        vals = grupe_data[grupa]
        g_curr = vals[last_idx]
        g_prev = vals[last_idx - 1] if num_s >= 2 else 0
        g_avg6 = round(sum(vals[last_idx - n6:last_idx]) / n6) if n6 > 0 else 0
        g_yoy = vals[yoy_idx] if yoy_idx is not None else None
        g_pct_mom, g_pct_6m = pct(g_curr, g_prev), pct(g_curr, g_avg6)
        g_pct_yoy = pct(g_curr, g_yoy) if g_yoy else None
        udeo = round(g_curr / curr_total * 100, 1) if curr_total > 0 else 0
        colors = grupa_boje[grupa]
        safe_id = grupa.replace(' ', '_').replace('.', '')
        gp = [f'<b style="color:{colors["main"]}">{esc(grupa)}</b> beleži']
        if num_s >= 2:
            gp.append(f'{badge(g_pct_mom)} u odnosu na prethodni mesec '
                      f'<span style="color:#8b90a5">({prev_naziv}: {fmtnum(g_prev)})</span>,')
        if n6 > 0:
            gp.append(f'{badge(g_pct_6m)} u odnosu na 6M prosek '
                      f'<span style="color:#8b90a5">({fmtnum(g_avg6)})</span>')
        if g_pct_yoy is not None:
            gp.append(f'i {badge(g_pct_yoy)} YoY u odnosu na {yoy_naziv} '
                      f'<span style="color:#8b90a5">({fmtnum(g_yoy)})</span>')
        grupe_chart_data_js[safe_id] = {
            'labels': [f'{poslednji_naziv} (trenutno)', f'{prev_naziv} (-1M)', 'Ø 6 meseci',
                       f'{yoy_naziv} (YoY)' if g_yoy else ''],
            'values': [g_curr, g_prev, g_avg6, g_yoy if g_yoy else 0],
            'has_yoy': g_yoy is not None, 'main_color': colors['main']}
        grupe_blocks_html += f'''
<div class="grupa-block">
  <div class="grupa-header">
    <div class="grupa-title-wrap">
      <span class="grupa-pill" style="background:{colors['bg']};color:{colors['main']}">{esc(grupa)}</span>
      <span class="grupa-udeo">{udeo}% ukupnog miksa</span>
    </div>
    <span class="grupa-kom" style="color:{colors['main']}">{fmtnum(g_curr)} kom</span>
  </div>
  <p class="grupa-text">{' '.join(gp) + '.'}</p>
  <div class="grupa-chart-wrap"><canvas id="chart-{safe_id}"></canvas></div>
</div>
'''

    mini1 = ({'labels': [prev_naziv, poslednji_naziv], 'values': [prev_total, curr_total],
              'color': '#dc2626' if pct_mom < 0 else '#16a34a'} if num_s >= 2 else None)
    mini2 = ({'labels': ['Ø 6M', poslednji_naziv], 'values': [avg6_total, curr_total],
              'color': '#dc2626' if pct_6m < 0 else '#16a34a'} if n6 > 0 else None)
    mini3 = ({'labels': [yoy_naziv, poslednji_naziv], 'values': [yoy_total, curr_total],
              'color': '#dc2626' if (pct_yoy or 0) < 0 else '#16a34a'}
             if pct_yoy is not None else None)
    analitika_data_json = json.dumps({
        'big': {'labels': nazivi, 'values': grand_totals, 'curr_idx': last_idx,
                'yoy_idx': yoy_idx if yoy_idx is not None else -1},
        'mini1': mini1, 'mini2': mini2, 'mini3': mini3, 'grupe': grupe_chart_data_js},
        ensure_ascii=False)

    def mini_card(label, pct_val, diff_val, sub_left, sub_right, chart_id):
        if pct_val is None:
            return ''
        color = '#16a34a' if pct_val >= 0 else '#dc2626'
        sign = '+' if pct_val >= 0 else ''
        return f'''<div class="mini-card" style="border-top:3px solid {color}">
  <div class="mc-lbl">{label}</div>
  <div class="mc-row"><span class="mc-pct" style="color:{color}">{sign}{pct_val}%</span>
    <span class="mc-diff">{fmt_diff(diff_val)} kom</span></div>
  <div class="mc-chart"><canvas id="{chart_id}"></canvas></div>
  <div class="mc-sub"><span>{sub_left}</span>
    <span style="color:{color};font-weight:700">{sub_right}</span></div>
</div>'''

    mini1_html = (mini_card('vs PRETHODNI MESEC', pct_mom, diff_mom,
                            f'{prev_naziv}: {fmtnum(prev_total)}',
                            f'{poslednji_naziv}: {fmtnum(curr_total)}', 'mini1')
                  if num_s >= 2 else '')
    mini2_html = (mini_card('vs 6M PROSEK', pct_6m, diff_6m, f'Ø 6M: {fmtnum(avg6_total)}',
                            f'{poslednji_naziv}: {fmtnum(curr_total)}', 'mini2') if n6 > 0 else '')
    mini3_html = (mini_card(f'YoY ({yoy_naziv} → {poslednji_naziv})', pct_yoy, diff_yoy,
                            f'{yoy_naziv}: {fmtnum(yoy_total) if yoy_total else "—"}',
                            f'{poslednji_naziv}: {fmtnum(curr_total)}', 'mini3')
                  if pct_yoy is not None else '')

    # ---------- top/bottom + alarmi ----------
    sistem_monthly = {}
    for sistem in sistemi_lista:
        s_data = df_clean[df_clean['SISTEM'] == sistem]
        vals = []
        for g, m in periodi:
            mask = (s_data['Godina'] == int(g)) & (s_data['Mesec'] == int(m))
            vals.append(round(float(s_data.loc[mask, 'Prodata kolicina ka krajnjem kupcu'].sum())))
        sistem_monthly[sistem] = vals

    sistem_mom_changes = []
    for sistem in sistemi_lista:
        vals = sistem_monthly[sistem]
        curr = vals[last_idx]
        prev = vals[last_idx - 1] if num_s >= 2 else 0
        if prev > 0 and curr > 0:
            sistem_mom_changes.append({'sistem': sistem, 'pct': pct(curr, prev),
                                       'abs': curr - prev, 'curr': curr, 'prev': prev})
    top3 = sorted([s for s in sistem_mom_changes if s['pct'] > 0], key=lambda x: -x['pct'])[:3]
    bottom3 = sorted([s for s in sistem_mom_changes if s['pct'] < 0], key=lambda x: x['pct'])[:3]

    alarmi = []
    for sistem in sistemi_lista:
        vals = sistem_monthly[sistem]
        curr = vals[last_idx]
        if num_s >= 4:
            consec = 0
            for i in range(last_idx, 0, -1):
                if vals[i] < vals[i - 1] and vals[i - 1] > 0:
                    consec += 1
                else:
                    break
            if consec >= 3:
                alarmi.append({'level': 'KRITIČNO', 'color': '#dc2626', 'bg': '#fef2f2',
                               'naziv': sistem, 'tip': 'sistem',
                               'tekst': f'pad {consec} meseca uzastopno '
                                        f'({pct(vals[last_idx], vals[last_idx - consec]):+}% kumulativno)'})
        if curr == 0 and any(v > 0 for v in vals[:-1]):
            zadnji = None
            for i in range(last_idx - 1, -1, -1):
                if vals[i] > 0:
                    zadnji = i
                    break
            if zadnji is not None:
                alarmi.append({'level': 'UPOZORENJE', 'color': '#92400e', 'bg': '#fffbeb',
                               'naziv': sistem, 'tip': 'sistem',
                               'tekst': f'nema prodaje u {poslednji_naziv} (poslednja prodaja: '
                                        f'{nazivi[zadnji]} — {fmtnum(vals[zadnji])} kom)'})
        if curr > 0 and n6 > 0:
            avg6 = round(sum(vals[last_idx - n6:last_idx]) / n6)
            if avg6 > 0 and pct(curr, avg6) <= -15:
                alarmi.append({'level': 'UPOZORENJE', 'color': '#92400e', 'bg': '#fffbeb',
                               'naziv': sistem, 'tip': 'sistem',
                               'tekst': f'prodaja ispod 6M proseka za {abs(pct(curr, avg6))}% '
                                        f'({fmtnum(curr)} vs Ø {fmtnum(avg6)})'})
        if curr > 0 and curr == max(vals) and len(vals) > 1 and curr > sorted(vals)[-2]:
            alarmi.append({'level': 'REKORD', 'color': '#16a34a', 'bg': 'rgba(22,163,74,0.06)',
                           'naziv': sistem, 'tip': 'sistem',
                           'tekst': f'rekord prodaje u {poslednji_naziv} ({fmtnum(curr)} kom)'})

    for sistem in sistemi_lista:
        s_data = df_clean[df_clean['SISTEM'] == sistem]
        for grupa in TRAZENE_GRUPE:
            g_data = s_data[s_data['Grupa artikla'].astype(str).str.strip() == grupa]
            sg = []
            for g, m in periodi:
                mask = (g_data['Godina'] == int(g)) & (g_data['Mesec'] == int(m))
                sg.append(round(float(g_data.loc[mask, 'Prodata kolicina ka krajnjem kupcu'].sum())))
            if sum(sg) == 0:
                continue
            if num_s >= 4:
                consec = 0
                for i in range(last_idx, 0, -1):
                    if sg[i] < sg[i - 1] and sg[i - 1] > 0:
                        consec += 1
                    else:
                        break
                if consec >= 3:
                    alarmi.append({'level': 'KRITIČNO', 'color': '#dc2626', 'bg': '#fef2f2',
                                   'naziv': f'{sistem} / {grupa}', 'tip': 'kombinacija',
                                   'tekst': f'pad {consec} meseca uzastopno '
                                            f'({pct(sg[last_idx], sg[last_idx - consec]):+}% kumulativno)'})
            if sg[last_idx] == 0 and num_s >= 4:
                recent = sum(sg[max(0, last_idx - 3):last_idx]) / min(3, last_idx)
                if recent >= 50:
                    alarmi.append({'level': 'UPOZORENJE', 'color': '#92400e', 'bg': '#fffbeb',
                                   'naziv': f'{sistem} / {grupa}', 'tip': 'kombinacija',
                                   'tekst': f'nema prodaje u {poslednji_naziv} '
                                            f'(prethodno Ø {fmtnum(round(recent))} kom/mesec)'})

    level_order = {'KRITIČNO': 0, 'UPOZORENJE': 1, 'REKORD': 2}
    alarmi.sort(key=lambda x: (level_order.get(x['level'], 99), x['naziv']))
    bez_prodaje = {a['naziv'] for a in alarmi if a['tip'] == 'sistem' and 'nema prodaje' in a['tekst']}
    sa_padom = {a['naziv'] for a in alarmi
                if a['tip'] == 'sistem' and 'pad' in a['tekst'] and 'uzastopno' in a['tekst']}
    _f = []
    for a in alarmi:
        if a['tip'] == 'kombinacija':
            sis0 = a['naziv'].split(' / ')[0]
            if 'nema prodaje' in a['tekst'] and sis0 in bez_prodaje:
                continue
            if 'pad' in a['tekst'] and 'uzastopno' in a['tekst'] and sis0 in sa_padom:
                continue
        _f.append(a)
    alarmi = _f

    def top_bottom_row(item, color):
        return (f'<div class="tb-row"><span class="tb-name">{esc(item["sistem"])}</span>'
                f'<span class="tb-vals"><b style="color:{color}">'
                f'{"+" if item["pct"] >= 0 else ""}{item["pct"]}%</b> '
                f'<span class="tb-diff">({fmt_diff(item["abs"])})</span></span></div>')

    top_rows_html = ('\n'.join(top_bottom_row(s, '#16a34a') for s in top3)
                     if top3 else '<div class="tb-empty">nema sistema sa rastom</div>')
    bottom_rows_html = ('\n'.join(top_bottom_row(s, '#dc2626') for s in bottom3)
                        if bottom3 else '<div class="tb-empty">nema sistema sa padom</div>')
    if alarmi:
        alarmi_html = ''
        for a in alarmi:
            alarmi_html += (f'<div class="alarm-row" style="background:{a["bg"]};'
                            f'border-left:3px solid {a["color"]}"><div class="alarm-text">'
                            f'<b style="color:{a["color"]};font-family:\'IBM Plex Mono\',monospace">'
                            f'{esc(a["naziv"])}</b> <span style="color:var(--t2)">— '
                            f'{esc(a["tekst"])}</span></div>'
                            f'<span class="alarm-level" style="color:{a["color"]}">'
                            f'{a["level"]}</span></div>')
    else:
        alarmi_html = '<div class="alarm-empty">Nema detektovanih alarma — sve u redu.</div>'
    n_kritic = len([a for a in alarmi if a['level'] == 'KRITIČNO'])
    n_upoz = len([a for a in alarmi if a['level'] == 'UPOZORENJE'])
    n_rekord = len([a for a in alarmi if a['level'] == 'REKORD'])

    top_bottom_html = f'''
<div class="tb-card">
  <div class="tb-header">
    <h3>🎯 TOP / BOTTOM SISTEMI — {poslednji_naziv.upper()} vs {prev_naziv.upper() if prev_naziv else "—"}</h3>
    <span class="tb-sub">najveći rast i pad u poslednjem mesecu</span>
  </div>
  <div class="tb-grid">
    <div class="tb-col tb-col-up"><div class="tb-col-title" style="color:#16a34a">▲ NAJVEĆI RAST</div>
      {top_rows_html}</div>
    <div class="tb-col tb-col-down"><div class="tb-col-title" style="color:#dc2626">▼ NAJVEĆI PAD</div>
      {bottom_rows_html}</div>
  </div>
</div>
<div class="alarm-card">
  <div class="alarm-header">
    <h3>⚠️ ALARMI — sistemi i grupe koje zahtevaju pažnju</h3>
    <div class="alarm-counts">
      {'<span class="ac-kritic">●' + str(n_kritic) + ' kritično</span>' if n_kritic else ''}
      {'<span class="ac-upoz">●' + str(n_upoz) + ' upozorenja</span>' if n_upoz else ''}
      {'<span class="ac-rekord">●' + str(n_rekord) + ' rekorda</span>' if n_rekord else ''}
    </div>
  </div>
  <div class="alarm-list">{alarmi_html}</div>
</div>
'''

    analitika_html = f'''
<div class="analitika-wrap">
  <div class="analitika-headline">
    <div class="ah-top"><h2>📊 ANALIZA PRODAJE — {poslednji_naziv.upper()}</h2>
      <span class="ah-badge">automatski ažurirano</span></div>
    <p class="ah-text">{headline_text}</p>
  </div>
  <div class="mini-cards-wrap">{mini1_html}{mini2_html}{mini3_html}</div>
  <div class="big-chart-card">
    <div class="bc-header"><h3>📈 TREND PRODAJE — {num_s} MESECI</h3>
      <div class="bc-legend">
        <span><span class="lg-line" style="background:#2563eb"></span>Mesečna prodaja</span>
        <span><span class="lg-dash"></span>6M prosek</span>
        <span><span class="lg-dot"></span>Isti mesec prethodne godine</span>
      </div>
    </div>
    <div class="bc-chart"><canvas id="bigTrendChart"></canvas></div>
  </div>
  <div class="grupe-section-header">
    <h2>📦 ANALIZA PO GRUPAMA — {poslednji_naziv.upper()}</h2>
    <div class="gsh-sub">Detaljna razrada za HQD 1000, NERD 2000 i SYX</div>
  </div>
  {grupe_blocks_html}
  {top_bottom_html}
</div>
'''

    # ---------- profitabilnost ----------
    profit_rows, si, profit_grand_totals = [], 0, [0] * num_p
    for sistem in sistemi_lista:
        si += 1
        sid2 = f"pr{si}"
        pd_s = profit_data_js[sistem]
        status = mapa_placanja.get(sistem, None)
        is_f = str(status) in ['1', '1.0']
        is_o = str(status) in ['0', '0.0']
        profit_niz, mkt_niz, dod_niz = pd_s['profit'], pd_s['mkt'], pd_s['dod']
        uk_niz = list(mkt_niz)
        for tid in trosak_ids_list[1:]:
            src = dod_niz if tid == 'dod' else pd_s.get(tid, [0] * num_p)
            for j in range(num_p):
                uk_niz[j] += src[j]
        neto_niz = [profit_niz[j] - uk_niz[j] for j in range(num_p)]
        for j in range(num_p):
            profit_grand_totals[j] += neto_niz[j]
        nacin = "po fakturi" if is_f else ("po odjavi" if is_o else "—")
        ncls = "nf" if is_f else "no"
        r = (f'<tr class="sr profit-sr" data-sistem="{esc(sistem)}" data-sid="{sid2}" '
             f'onclick="tog(\'{sid2}\')" style="cursor:pointer"><td>'
             f'<button class="be" id="b-{sid2}">+</button><span class="sn">{esc(sistem)}</span></td>'
             f'<td class="{ncls}">{nacin}</td><td class="nb">Neto profit</td>')
        for v in neto_niz:
            r += f'<td class="nb {"np" if v >= 0 else "nn"}">{fmtnum(v)}</td>'
        neto_total = sum(neto_niz)
        ncl = "np" if neto_total >= 0 else "nn"
        r += f'<td class="nb {ncl}" style="font-size:13px">{fmtnum(neto_total)}</td></tr>'
        profit_rows.append(r)
        stavka = "Profit (promet)" if is_f else ("Profit (odjava)" if is_o else "—")
        r = (f'<tr class="cr hidden" data-p="{sid2}" data-sistem="{esc(sistem)}"><td></td><td></td>'
             f'<td class="cl" style="color:var(--t2);font-style:normal;font-weight:600">{stavka}</td>')
        for v in profit_niz:
            r += f'<td class="n" style="font-weight:600;color:var(--t2)">{fmtnum(v)}</td>'
        r += (f'<td class="n" style="font-weight:700;color:var(--t2)">'
              f'{fmtnum(sum(profit_niz))}</td></tr>')
        profit_rows.append(r)
        r = (f'<tr class="cr hidden" data-p="{sid2}" data-sistem="{esc(sistem)}" data-trosak="mkt">'
             f'<td></td><td></td><td class="cl">Troškovi marketinga</td>')
        for v in mkt_niz:
            r += f'<td class="n cc">{fmtnum(v)}</td>'
        r += f'<td class="n cc" style="font-weight:700">{fmtnum(sum(mkt_niz))}</td></tr>'
        profit_rows.append(r)
        for ki, kat in enumerate(trosak_kolone):
            tid = f"t{ki}"
            vals = pd_s.get(tid, [0] * num_p)
            r = (f'<tr class="cr hidden" data-p="{sid2}" data-sistem="{esc(sistem)}" '
                 f'data-trosak="{tid}"><td></td><td></td>'
                 f'<td class="cl">{esc(TROSAK_NAZIVI.get(kat, kat))}</td>')
            for v in vals:
                r += f'<td class="n cc">{fmtnum(v)}</td>'
            r += f'<td class="n cc" style="font-weight:700">{fmtnum(sum(vals))}</td></tr>'
            profit_rows.append(r)
        r = (f'<tr class="cr hidden" data-p="{sid2}" data-sistem="{esc(sistem)}" data-trosak="dod">'
             f'<td></td><td></td><td class="cl">Dodatni mesečni trošak</td>')
        for v in dod_niz:
            r += f'<td class="n cc">{fmtnum(v)}</td>'
        r += f'<td class="n cc" style="font-weight:700">{fmtnum(sum(dod_niz))}</td></tr>'
        profit_rows.append(r)
        r = (f'<tr class="ctr hidden ukupni-row" data-p="{sid2}" data-sistem="{esc(sistem)}" '
             f'data-sid="{sid2}"><td></td><td></td><td class="ctl">UKUPNI TROŠKOVI</td>')
        for v in uk_niz:
            r += f'<td class="n ctc">{fmtnum(v)}</td>'
        r += f'<td class="n ctc" style="font-size:12px">{fmtnum(sum(uk_niz))}</td></tr>'
        profit_rows.append(r)
        r = (f'<tr class="nr hidden neto-row" data-p="{sid2}" data-sistem="{esc(sistem)}" '
             f'data-sid="{sid2}"><td></td><td></td><td class="nl">NETO PROFIT</td>')
        for v in neto_niz:
            r += f'<td class="n {"np" if v >= 0 else "nn"}">{fmtnum(v)}</td>'
        r += f'<td class="n {ncl}" style="font-size:13px">{fmtnum(neto_total)}</td></tr>'
        profit_rows.append(r)
        profit_rows.append(f'<tr class="sep" data-sistem="{esc(sistem)}"><td colspan="999"></td></tr>')
    pgt = sum(profit_grand_totals)
    tr_p = '<tr class="totalrow" id="profit-total"><td class="total-label">TOTAL</td><td></td><td></td>'
    for v in profit_grand_totals:
        tr_p += f'<td class="nb total-cell {"np" if v >= 0 else "nn"}">{fmtnum(v)}</td>'
    tr_p += (f'<td class="nb total-cell {"np" if pgt >= 0 else "nn"}" style="font-size:13px">'
             f'{fmtnum(pgt)}</td></tr>')
    profit_rows.append(tr_p)

    # ---------- uspešnost akcije (tabela po odjavi) ----------
    drv_rows, si, drv_grand_totals = [], 0, [0] * num_d
    for sistem in sistemi_lista:
        si += 1
        sid3 = f"dv{si}"
        dv = drv_data_js[sistem]
        profit_niz, mkt_niz, dod_niz = dv['profit'], dv['mkt'], dv['dod']
        uk_niz = list(mkt_niz)
        for tid in trosak_ids_list[1:]:
            src = dod_niz if tid == 'dod' else dv.get(tid, [0] * num_d)
            for j in range(num_d):
                uk_niz[j] += src[j]
        neto_niz = [profit_niz[j] - uk_niz[j] for j in range(num_d)]
        for j in range(num_d):
            drv_grand_totals[j] += neto_niz[j]
        neto_total = sum(neto_niz)
        ncl = "np" if neto_total >= 0 else "nn"
        r = (f'<tr class="sr drv-sr" data-sistem="{esc(sistem)}" data-sid="{sid3}" '
             f'onclick="tog(\'{sid3}\')" style="cursor:pointer"><td>'
             f'<button class="be" id="b-{sid3}">+</button><span class="sn">{esc(sistem)}</span></td>'
             f'<td class="no">po odjavi</td><td class="nb">Neto profit</td>')
        for v in neto_niz:
            r += f'<td class="nb {"np" if v >= 0 else "nn"}">{fmtnum(v)}</td>'
        r += f'<td class="nb {ncl}" style="font-size:13px">{fmtnum(neto_total)}</td></tr>'
        drv_rows.append(r)
        r = (f'<tr class="cr hidden" data-p="{sid3}" data-sistem="{esc(sistem)}"><td></td><td></td>'
             f'<td class="cl" style="color:var(--t2);font-style:normal;font-weight:600">'
             f'Profit (odjava)</td>')
        for v in profit_niz:
            r += f'<td class="n" style="font-weight:600;color:var(--t2)">{fmtnum(v)}</td>'
        r += (f'<td class="n" style="font-weight:700;color:var(--t2)">'
              f'{fmtnum(sum(profit_niz))}</td></tr>')
        drv_rows.append(r)
        r = (f'<tr class="cr hidden" data-p="{sid3}" data-sistem="{esc(sistem)}" data-trosak="mkt">'
             f'<td></td><td></td><td class="cl">Troškovi marketinga</td>')
        for v in mkt_niz:
            r += f'<td class="n cc">{fmtnum(v)}</td>'
        r += f'<td class="n cc" style="font-weight:700">{fmtnum(sum(mkt_niz))}</td></tr>'
        drv_rows.append(r)
        for ki, kat in enumerate(trosak_kolone):
            tid = f"t{ki}"
            vals = dv.get(tid, [0] * num_d)
            r = (f'<tr class="cr hidden" data-p="{sid3}" data-sistem="{esc(sistem)}" '
                 f'data-trosak="{tid}"><td></td><td></td>'
                 f'<td class="cl">{esc(TROSAK_NAZIVI.get(kat, kat))}</td>')
            for v in vals:
                r += f'<td class="n cc">{fmtnum(v)}</td>'
            r += f'<td class="n cc" style="font-weight:700">{fmtnum(sum(vals))}</td></tr>'
            drv_rows.append(r)
        r = (f'<tr class="cr hidden" data-p="{sid3}" data-sistem="{esc(sistem)}" data-trosak="dod">'
             f'<td></td><td></td><td class="cl">Dodatni mesečni trošak</td>')
        for v in dod_niz:
            r += f'<td class="n cc">{fmtnum(v)}</td>'
        r += f'<td class="n cc" style="font-weight:700">{fmtnum(sum(dod_niz))}</td></tr>'
        drv_rows.append(r)
        r = (f'<tr class="ctr hidden ukupni-row-dv" data-p="{sid3}" data-sistem="{esc(sistem)}" '
             f'data-sid="{sid3}"><td></td><td></td><td class="ctl">UKUPNI TROŠKOVI</td>')
        for v in uk_niz:
            r += f'<td class="n ctc">{fmtnum(v)}</td>'
        r += f'<td class="n ctc" style="font-size:12px">{fmtnum(sum(uk_niz))}</td></tr>'
        drv_rows.append(r)
        r = (f'<tr class="nr hidden neto-row-dv" data-p="{sid3}" data-sistem="{esc(sistem)}" '
             f'data-sid="{sid3}"><td></td><td></td><td class="nl">NETO PROFIT</td>')
        for v in neto_niz:
            r += f'<td class="n {"np" if v >= 0 else "nn"}">{fmtnum(v)}</td>'
        r += f'<td class="n {ncl}" style="font-size:13px">{fmtnum(neto_total)}</td></tr>'
        drv_rows.append(r)
        drv_rows.append(f'<tr class="sep" data-sistem="{esc(sistem)}"><td colspan="999"></td></tr>')
    dgt = sum(drv_grand_totals)
    tr_d = '<tr class="totalrow" id="drv-total"><td class="total-label">TOTAL</td><td></td><td></td>'
    for v in drv_grand_totals:
        tr_d += f'<td class="nb total-cell {"np" if v >= 0 else "nn"}">{fmtnum(v)}</td>'
    tr_d += (f'<td class="nb total-cell {"np" if dgt >= 0 else "nn"}" style="font-size:13px">'
             f'{fmtnum(dgt)}</td></tr>')
    drv_rows.append(tr_d)

    # ---------- analiza uspešnosti akcije ----------
    df_art_full = df_clean[df_clean['Artikl'].notna()].copy()
    df_art_full['ima_akciju'] = df_art_full['AKCIJSKE CENE'].fillna(0) > 0
    akcija_blokovi_data = []

    def bruto_profit_artikli_mesec(s_art_df, artikli, godina, mesec):
        mask = ((s_art_df['Godina'] == int(godina)) & (s_art_df['Mesec'] == int(mesec))
                & (s_art_df['Artikl'].isin(artikli)))
        if not mask.any():
            return 0
        return float(s_art_df.loc[mask, 'Profit'].sum())

    for sistem in sistemi_lista:
        s_art_all = df_art_full[df_art_full['SISTEM'] == sistem].copy()
        if len(s_art_all) == 0:
            continue
        s_art_all['Grupa_clean'] = s_art_all['Grupa artikla'].astype(str).str.strip()
        s_art = s_art_all[s_art_all['Grupa_clean'].isin(POSMATRANE_GRUPE)].copy()
        if len(s_art) == 0:
            continue
        g_per_mes = s_art.groupby(['Godina', 'Mesec']).agg(na_akciji=('ima_akciju', 'sum')).reset_index()
        g_per_mes['ima_neku_akciju'] = g_per_mes['na_akciji'] > 0
        g_per_mes = g_per_mes.sort_values(['Godina', 'Mesec']).reset_index(drop=True)
        akc_df = g_per_mes[g_per_mes['ima_neku_akciju']]
        red_df = g_per_mes[~g_per_mes['ima_neku_akciju']]
        if len(akc_df) == 0 or len(red_df) == 0:
            continue
        last_akc_g, last_akc_m = int(akc_df.iloc[-1]['Godina']), int(akc_df.iloc[-1]['Mesec'])
        last_neakc_g, last_neakc_m = int(red_df.iloc[-1]['Godina']), int(red_df.iloc[-1]['Mesec'])
        akc_data = s_art[(s_art['Godina'] == last_akc_g) & (s_art['Mesec'] == last_akc_m)
                         & (s_art['ima_akciju'])]
        if len(akc_data) == 0:
            continue
        akcijski_artikli = akc_data['Artikl'].unique().tolist()
        neakc_data = s_art[(s_art['Godina'] == last_neakc_g) & (s_art['Mesec'] == last_neakc_m)
                           & (s_art['Artikl'].isin(akcijski_artikli))]
        grupe_breakdown, komparabilne_grupe = [], []
        total_akc_qty = total_neakc_qty = 0
        for grupa in sorted(akc_data['Grupa artikla'].dropna().astype(str).str.strip().unique()):
            akc_g = akc_data[akc_data['Grupa artikla'].astype(str).str.strip() == grupa]
            neakc_g = neakc_data[neakc_data['Grupa artikla'].astype(str).str.strip() == grupa]
            akc_g_cena = float(akc_g['AKCIJSKE CENE'].mean()) if len(akc_g) else 0
            ref = float(akc_g['Redovna MP CENA'].mean()) if len(akc_g) else 0
            stv = float(neakc_g['Redovna MP CENA'].mean()) if len(neakc_g) else 0
            neakc_g_cena = stv if stv > 0 else ref
            akc_g_qty = float(akc_g['Prodata kolicina ka krajnjem kupcu'].sum())
            neakc_g_qty = (float(neakc_g['Prodata kolicina ka krajnjem kupcu'].sum())
                           if len(neakc_g) else 0)
            has = len(neakc_g) > 0
            if has:
                total_akc_qty += akc_g_qty
                total_neakc_qty += neakc_g_qty
                komparabilne_grupe.append(grupa)
            grupe_breakdown.append({'grupa': grupa, 'akc_cena': akc_g_cena,
                                    'neakc_cena': neakc_g_cena, 'akc_qty': akc_g_qty,
                                    'neakc_qty': neakc_g_qty, 'has_neakc_data': has,
                                    'broj_artikala': len(akc_g)})
        if not komparabilne_grupe:
            continue
        art_komp = akc_data[akc_data['Grupa artikla'].astype(str).str.strip()
                            .isin(komparabilne_grupe)]['Artikl'].unique().tolist()
        akc_bruto = bruto_profit_artikli_mesec(s_art, art_komp, last_akc_g, last_akc_m)
        neakc_bruto = bruto_profit_artikli_mesec(s_art, art_komp, last_neakc_g, last_neakc_m)
        pct_qty_total = pct(total_akc_qty, total_neakc_qty) if total_neakc_qty > 0 else None
        pct_bruto = pct(akc_bruto, neakc_bruto) if neakc_bruto != 0 else None
        if pct_bruto is None:
            status_label, status_color, status_bg = '— nema reference', '#8b90a5', 'rgba(139,144,165,0.08)'
            zb, zg, zc = '#8b90a5', 'rgba(139,144,165,0.05),rgba(0,0,0,0.02)', '#5a5f7a'
        elif pct_bruto >= 10:
            status_label, status_color, status_bg = '✓ ISPLATILA SE', '#16a34a', 'rgba(22,163,74,0.08)'
            zb, zg, zc = '#16a34a', 'rgba(22,163,74,0.05),rgba(37,99,235,0.03)', '#16a34a'
        elif pct_bruto >= -10:
            status_label, status_color, status_bg = '~ MARGINALNO', '#92400e', 'rgba(245,158,11,0.08)'
            zb, zg, zc = '#f59e0b', 'rgba(245,158,11,0.05),rgba(220,38,38,0.03)', '#92400e'
        else:
            status_label, status_color, status_bg = '✗ NIJE SE ISPLATILA', '#dc2626', 'rgba(220,38,38,0.08)'
            zb, zg, zc = '#dc2626', 'rgba(220,38,38,0.05),rgba(245,158,11,0.03)', '#dc2626'
        monthly_qty, monthly_profit, monthly_is_akcija = [], [], []
        for gper, mper in periodi:
            m_data = s_art[(s_art['Godina'] == int(gper)) & (s_art['Mesec'] == int(mper))
                           & (s_art['Artikl'].isin(art_komp))]
            monthly_qty.append(round(float(m_data['Prodata kolicina ka krajnjem kupcu'].sum())))
            monthly_profit.append(round(float(m_data['Profit'].sum())))
            monthly_is_akcija.append(bool((m_data['ima_akciju'] == True).any()))
        akcija_blokovi_data.append({
            'sistem': sistem, 'broj_artikala': len(art_komp),
            'last_akc_naziv': f"{MAPA_MESECI[last_akc_m]} {str(last_akc_g)[-2:]}",
            'last_neakc_naziv': f"{MAPA_MESECI[last_neakc_m]} {str(last_neakc_g)[-2:]}",
            'grupe_breakdown': grupe_breakdown, 'total_akc_qty': total_akc_qty,
            'total_neakc_qty': total_neakc_qty, 'pct_qty_total': pct_qty_total,
            'akc_bruto': akc_bruto, 'neakc_bruto': neakc_bruto, 'pct_bruto': pct_bruto,
            'status_label': status_label, 'status_color': status_color, 'status_bg': status_bg,
            'zaklj_border': zb, 'zaklj_bg_grad': zg, 'zaklj_color': zc,
            'monthly_qty': monthly_qty, 'monthly_profit': monthly_profit,
            'monthly_is_akcija': monthly_is_akcija})

    akcija_blokovi_data.sort(key=lambda x: 9999 if x['pct_bruto'] is None else -x['pct_bruto'])

    akcija_html_blokovi = ''
    for d in akcija_blokovi_data:
        safe_id = d['sistem'].replace(' ', '_').replace('/', '_').replace('.', '')
        grupe_rows_parts = []
        for gb in d['grupe_breakdown']:
            pct_cena_g = pct(gb['akc_cena'], gb['neakc_cena']) if gb['neakc_cena'] > 0 else None
            pct_qty_g = pct(gb['akc_qty'], gb['neakc_qty']) if gb['neakc_qty'] > 0 else None
            if gb['has_neakc_data']:
                qty_neakc_cell = fmtnum(gb['neakc_qty'])
                pct_qty_cell = (f'<span style="color:{color_pct(pct_qty_g)};font-weight:700">'
                                f'{fmt_pct_signed(pct_qty_g)}</span>')
            else:
                qty_neakc_cell = '<span style="color:var(--t3);font-style:italic">nema</span>'
                pct_qty_cell = '<span style="color:var(--t3)">—</span>'
            grupe_rows_parts.append(
                '<tr>'
                f'<td class="g-naziv">{esc(gb["grupa"])} '
                f'<span class="g-broj">({gb["broj_artikala"]} art.)</span></td>'
                f'<td class="g-num">{fmtnum(gb["akc_cena"])} RSD</td>'
                f'<td class="g-num">{fmtnum(gb["neakc_cena"])} RSD</td>'
                f'<td class="g-num" style="color:{color_pct(pct_cena_g)}">'
                f'{fmt_pct_signed(pct_cena_g)}</td>'
                f'<td class="g-num" style="color:var(--grn);font-weight:700">'
                f'{fmtnum(gb["akc_qty"])}</td>'
                f'<td class="g-num">{qty_neakc_cell}</td>'
                f'<td class="g-num">{pct_qty_cell}</td></tr>')
        if d['pct_bruto'] is None:
            zakljucak = (f'Za <b>{d["broj_artikala"]} artikala</b> koji su bili na akciji '
                         f'({esc(d["last_akc_naziv"])} vs {esc(d["last_neakc_naziv"])}): prodato je '
                         f'<b style="color:var(--grn)">{fmtnum(d["total_akc_qty"])} kom</b> u akciji. '
                         f'Nema dovoljno podataka iz neakcijskog meseca za poređenje profita.')
        else:
            if d['pct_bruto'] >= 10:
                komentar = "Lager očišćen efikasno, profit veći uprkos nižoj marži."
            elif d['pct_bruto'] >= -10:
                komentar = "Efekat neutralan — povećana prodaja kompenzuje sniženu maržu."
            else:
                komentar = ("Razmotri da li je popust preagresivan — gubitak marže veći "
                            "od dobitka u količini.")
            qty_part = ''
            if d['pct_qty_total'] is not None:
                qty_part = (f'Prodaja je <b style="color:{color_pct(d["pct_qty_total"])}">'
                            f'{fmt_pct_signed(d["pct_qty_total"])}</b> '
                            f'<span style="color:var(--t3)">({fmtnum(d["total_akc_qty"])} vs '
                            f'{fmtnum(d["total_neakc_qty"])} kom)</span>. ')
            zakljucak = (
                f'Za <b>{d["broj_artikala"]} artikala</b> koji su bili na akciji, poređenjem '
                f'<b>{esc(d["last_akc_naziv"])}</b> (akcija) i <b>{esc(d["last_neakc_naziv"])}</b> '
                f'(bez akcije): {qty_part}<br>'
                f'<span style="font-family:\'IBM Plex Mono\',monospace;font-size:12px">'
                f'Bruto profit (samo iz akcijskih grupa):</span> '
                f'<b style="color:{color_pct(d["pct_bruto"])};font-family:\'IBM Plex Mono\',monospace;'
                f'font-size:13px">{fmt_pct_signed(d["pct_bruto"])}</b> '
                f'<span style="color:var(--t3)">({fmtnum(d["akc_bruto"])} vs '
                f'{fmtnum(d["neakc_bruto"])} RSD)</span>. {komentar}')
        akcija_html_blokovi += (
            '<div class="ak-block"><div class="ak-block-header"><div class="ak-sistem-info">'
            f'<span class="ak-pill">{esc(d["sistem"])}</span>'
            f'<span class="ak-meta">{d["broj_artikala"]} artikala na akciji · '
            f'{esc(d["last_akc_naziv"])} (akcija) vs {esc(d["last_neakc_naziv"])} (redovno)</span></div>'
            f'<span class="ak-status" style="background:{d["status_bg"]};color:{d["status_color"]}">'
            f'{d["status_label"]}</span></div>'
            f'<div class="ak-zakljucak" style="background:linear-gradient(90deg,{d["zaklj_bg_grad"]});'
            f'border-left:3px solid {d["zaklj_border"]}">'
            f'<div class="ak-zakljucak-label" style="color:{d["zaklj_color"]}">📌 ZAKLJUČAK</div>'
            f'<p style="font-size:13px;line-height:1.7;color:var(--t1);margin:0">{zakljucak}</p></div>'
            '<div class="ak-grupe-table-wrap">'
            '<div class="ak-table-title">💰 Cene i količine — razloženo po grupama</div>'
            '<table class="ak-grupe-table"><thead><tr>'
            '<th rowspan="2" style="text-align:left">GRUPA</th>'
            '<th colspan="3" class="th-cena">CENA (RSD)</th>'
            '<th colspan="3" class="th-qty">KOLIČINA (kom)</th></tr>'
            '<tr><th class="th-akc">akcija</th><th class="th-neakc">redovno</th>'
            '<th class="th-raz">razlika</th><th class="th-akc">akcija</th>'
            '<th class="th-neakc">redovno</th><th class="th-raz">razlika</th></tr></thead>'
            f'<tbody>{"".join(grupe_rows_parts)}</tbody></table></div>'
            '<div class="ak-neto-row">'
            '<div class="ak-neto-box ak-neto-akc">'
            f'<div class="ak-neto-lbl">BRUTO PROFIT — {esc(d["last_akc_naziv"])} (akcija)</div>'
            f'<div class="ak-neto-val">{fmtnum(d["akc_bruto"])} RSD</div></div>'
            '<div class="ak-neto-box ak-neto-neakc">'
            f'<div class="ak-neto-lbl">BRUTO PROFIT — {esc(d["last_neakc_naziv"])} (redovno)</div>'
            f'<div class="ak-neto-val">{fmtnum(d["neakc_bruto"])} RSD</div></div>'
            '<div class="ak-neto-box ak-neto-razlika">'
            '<div class="ak-neto-lbl">RAZLIKA BRUTO PROFITA</div>'
            f'<div class="ak-neto-val" style="color:{color_pct(d["pct_bruto"])}">'
            f'{fmt_pct_signed(d["pct_bruto"])}</div></div></div>'
            '<div class="ak-chart-wrap"><div class="ak-chart-legend">'
            '<span class="ak-chart-title">📈 Prodaja akcijskih artikala + bruto profit po mesecu</span>'
            '<span class="ak-legend-items">'
            '<span><span class="ak-leg-sq ak-leg-akcija"></span>akcija</span>'
            '<span><span class="ak-leg-sq ak-leg-redovno"></span>redovno</span>'
            '<span><span class="ak-leg-profit"></span>bruto profit</span></span></div>'
            f'<div class="ak-chart"><canvas id="ak-chart-{safe_id}"></canvas></div></div></div>')

    akcija_chart_data = {}
    for d in akcija_blokovi_data:
        safe_id = d['sistem'].replace(' ', '_').replace('/', '_').replace('.', '')
        akcija_chart_data[safe_id] = {'labels': nazivi, 'values': d['monthly_qty'],
                                      'profit': d['monthly_profit'],
                                      'is_akcija': d['monthly_is_akcija']}
    akcija_data_json = json.dumps(akcija_chart_data, ensure_ascii=False)

    if akcija_blokovi_data:
        akcija_analiza_html = (
            '<div class="akcija-analiza-wrap"><div class="aa-header">'
            '<h2>🎯 ANALIZA USPEŠNOSTI AKCIJE — PO SISTEMIMA</h2>'
            '<div class="aa-sub">Poređenje poslednjeg akcijskog meseca i najbližeg neakcijskog '
            'meseca — samo za artikle koji su bili na akciji. Cene posebno po grupama. '
            'Sortirano po neto profitabilnosti.</div></div>' + akcija_html_blokovi + '</div>')
    else:
        akcija_analiza_html = ('<div style="padding:20px;text-align:center;color:var(--t3);'
                               'font-family:\'IBM Plex Mono\',monospace;font-size:12px">'
                               'Nema dovoljno podataka za analizu akcije.</div>')

    # ---------- obrt lagera ----------
    obrt_lagera_data = []
    for sistem in sistemi_lista:
        s_data = df_clean[df_clean['SISTEM'] == sistem]
        if len(s_data) == 0 or 'Zalihe' not in s_data.columns:
            continue
        s_zal = s_data.groupby(['Godina', 'Mesec']).agg(zal=('Zalihe', 'sum')).reset_index()
        s_zal = s_zal[s_zal['zal'] > 0].sort_values(['Godina', 'Mesec'])
        if len(s_zal) == 0:
            continue
        sys_last_g, sys_last_m = int(s_zal.iloc[-1]['Godina']), int(s_zal.iloc[-1]['Mesec'])
        trenutni_lager = float(s_zal.iloc[-1]['zal'])
        idx_in = None
        for i, (g, m) in enumerate(periodi):
            if int(g) == sys_last_g and int(m) == sys_last_m:
                idx_in = i
                break
        if idx_in is None:
            continue
        last_3_sys = periodi[max(0, idx_in - 2):idx_in + 1]
        prodaja_sum = 0
        for g3, m3 in last_3_sys:
            mask3 = (s_data['Godina'] == int(g3)) & (s_data['Mesec'] == int(m3))
            prodaja_sum += float(s_data.loc[mask3, 'Prodata kolicina ka krajnjem kupcu'].sum())
        prosek = prodaja_sum / len(last_3_sys) if last_3_sys else 0
        if prosek <= 0:
            continue
        obrt_lagera_data.append({'sistem': sistem, 'lager': trenutni_lager, 'prodaja_avg': prosek,
                                 'meseci': trenutni_lager / prosek,
                                 'last_mes_naziv': f"{MAPA_MESECI[sys_last_m]} {str(sys_last_g)[-2:]}"})
    obrt_lagera_data.sort(key=lambda x: x['meseci'])
    obrt_chart_json = json.dumps({
        'sistemi': [d['sistem'] for d in obrt_lagera_data],
        'meseci': [round(d['meseci'], 1) for d in obrt_lagera_data],
        'lager': [round(d['lager']) for d in obrt_lagera_data],
        'prodaja_avg': [round(d['prodaja_avg']) for d in obrt_lagera_data],
        'last_mes': [d['last_mes_naziv'] for d in obrt_lagera_data]}, ensure_ascii=False)
    obrt_lagera_html = ('''
<div class="obrt-wrap">
  <div class="obrt-header">
    <h3>⚙️ OBRT LAGERA PO SISTEMIMA</h3>
    <div class="obrt-sub">Trenutni lager (poslednji raspoloživi mesec po sistemu) / Ø prodaja
      poslednja 3 meseca = broj meseci da se obrne lager</div>
  </div>
  <div class="obrt-card">
    <div class="obrt-legend-row">
      <span class="obrt-title">📊 BROJ MESECI ZA OBRT LAGERA</span>
      <span class="obrt-legend">
        <span><span class="obrt-leg-sq" style="background:#16a34a"></span>≤3 mes (brz)</span>
        <span><span class="obrt-leg-sq" style="background:#0d9488"></span>3-6 mes</span>
        <span><span class="obrt-leg-sq" style="background:#f59e0b"></span>6-12 mes</span>
        <span><span class="obrt-leg-sq" style="background:#dc2626"></span>&gt;12 mes (spor)</span>
      </span>
    </div>
    <div class="obrt-chart"><canvas id="obrtChart"></canvas></div>
    <div class="obrt-note">Sortirano od najbržeg ka najsporijem obrtu. Manje meseci = bolje.</div>
  </div>
</div>''' if obrt_lagera_data else '')

    # ---------- zaglavlja tabela ----------
    ph = '<tr><th style="text-align:left;min-width:280px">SISTEM / GRUPA / ARTIKAL</th>'
    for i, n in enumerate(nazivi):
        ph += f'<th class="m{i}">{n}</th>'
    ph += '<th>TOTAL</th></tr>'
    prh = ('<tr><th style="text-align:left;min-width:200px">SISTEM</th><th>NAČIN</th>'
           '<th>STAVKA</th>' + ''.join(f'<th>{n}</th>' for n in nazivi_profit) + '<th>TOTAL</th></tr>')
    drvh = ('<tr><th style="text-align:left;min-width:200px">SISTEM</th><th>NAČIN</th>'
            '<th>STAVKA</th>' + ''.join(f'<th>{n}</th>' for n in nazivi_drv) + '<th>TOTAL</th></tr>')
    mh = ('<tr><th style="text-align:left;min-width:300px">ARTIKAL</th><th>GRUPA</th>'
          '<th>MAGACIN (kom)</th><th>Ø PRODAJA/mes</th><th>DANA ZALIHA</th><th>MESECI</th></tr>')
    zsh = ('<tr><th style="text-align:left;min-width:280px">SISTEM / GRUPA / ARTIKAL</th>'
           '<th>ZALIHE (kom)</th><th>PERIOD</th></tr>')

    info = f"{len(sistemi_lista)} sistema · {len(nazivi)} meseci"
    sistem_options = ''.join(f'<option value="{esc(s)}">{esc(s)}</option>' for s in sistemi_lista)
    godine_lista = sorted(set(int(g) for g, m in periodi))
    godina_options = ''.join(f'<option value="{g}">{g}</option>' for g in godine_lista)

    def _pop(wrap_id, pop_id, label, count_id, body, all_fn, wide=False):
        return (f'<span class="ms-wrap" id="{wrap_id}">'
                f'<button class="ms-btn" onclick="togglePop(event,\'{pop_id}\')">{label}: '
                f'<b id="{count_id}">–</b> ▾</button>'
                f'<div class="ms-pop{" ms-pop-wide" if wide else ""} hidden" id="{pop_id}">{body}'
                f'<div class="ms-actions">'
                f'<button class="mini-btn" onclick="{all_fn}(true)">Sve</button>'
                f'<button class="mini-btn" onclick="{all_fn}(false)">Nijedan</button>'
                f'<button class="mini-btn ms-close" onclick="togglePop(event,\'{pop_id}\')">'
                f'Zatvori</button></div></div></span>')

    sistem_checks = ''.join(
        f'<label class="mcb"><input type="checkbox" checked value="{esc(x)}" '
        f'onchange="applyFilters()"><span>{esc(x)}</span></label>' for x in sistemi_lista)
    grupa_checks = ''.join(
        f'<label class="mcb"><input type="checkbox" checked value="{esc(x)}" '
        f'onchange="applyFilters()"><span>{esc(x)}</span></label>' for x in sve_grupe)
    sistem_pop = _pop('sis-wrap', 'sis-pop', 'Sistemi', 'cnt-sistem',
                      f'<div class="ms-grid2" id="sistem-checks">{sistem_checks}</div>',
                      'allSistemi', wide=True)
    grupa_pop = _pop('gr-wrap', 'gr-pop', 'Grupe', 'cnt-grupa',
                     f'<div class="ms-grid2" id="grupa-checks">{grupa_checks}</div>',
                     'allGrupe', wide=True)

    mesec_checks = ''
    for god in godine_lista:
        idxs = [i for i, (g, m) in enumerate(periodi) if int(g) == god]
        if not idxs:
            continue
        kvadratici = ''.join(
            f'<label class="mcb"><input type="checkbox" checked value="{i}" data-god="{god}" '
            f'onchange="applyFilters()"><span>{MAPA_MESECI[int(periodi[i][1])]}</span></label>'
            for i in idxs)
        mesec_checks += (f'<div class="ms-god"><div class="ms-god-hdr">'
                         f'<span class="ms-god-nm">{god}</span><span>'
                         f'<button class="mini-btn" onclick="yearToggle({god},true)">sve</button>'
                         f'<button class="mini-btn" onclick="yearToggle({god},false)">ništa</button>'
                         f'</span></div><div class="ms-grid">{kvadratici}</div></div>')

    trosak_checks = ''.join(
        f'<label class="tcb"><input type="checkbox" checked value="{tid}" '
        f'onchange="applyProfitFilters()"><span>{tname}</span></label>'
        for tid, tname in zip(trosak_ids_list, trosak_names_list))
    trosak_checks_drv = ''.join(
        f'<label class="tcb"><input type="checkbox" checked value="{tid}" '
        f'onchange="applyDrvFilters()"><span>{tname}</span></label>'
        for tid, tname in zip(trosak_ids_list, trosak_names_list))

    prodaja_json = json.dumps(prodaja_data_js, ensure_ascii=False)
    profit_json = json.dumps(profit_data_js, ensure_ascii=False)
    drv_json = json.dumps(drv_data_js, ensure_ascii=False)

    # ---------- info trake ----------
    if promene_lista:
        pc_items = ''
        for p in promene_lista:
            c1 = get_color(p['stara'], p['grupa']) or '#e8ecf4'
            c2 = get_color(p['nova'], p['grupa']) or '#e8ecf4'
            strelica = '▼' if p['smer'] == 'dole' else ('▲' if p['smer'] == 'gore' else '=')
            boja_str = ('#dc2626' if p['smer'] == 'dole'
                        else ('#16a34a' if p['smer'] == 'gore' else '#8b90a5'))
            pc_items += (
                f'<div class="pc-item"><span class="pc-swatch" style="background:'
                f'linear-gradient(90deg,{c1} 0%,{c1} 50%,#1a1a2e 50%,#1a1a2e calc(50% + 2px),'
                f'{c2} calc(50% + 2px),{c2} 100%)"></span>'
                f'<span class="pc-sys">{esc(p["sistem"])}</span>'
                f'<span class="pc-gr">{esc(p["grupa"])}</span>'
                f'<span class="pc-per">{esc(p["mesec_naziv"])}</span>'
                f'<span class="pc-dan">od {p["dan"]}.{p["mesec"]}.</span>'
                f'<span class="pc-cene" style="color:{boja_str}">{fmtnum(p["stara"])} '
                f'{strelica} {fmtnum(p["nova"])} RSD</span></div>')
        promene_bar_html = (
            '<div class="pc-bar"><div class="pc-bar-title">⇄ PROMENE CENE U TOKU MESECA '
            f'<span class="pc-bar-sub">({len(promene_lista)}) — ćelija je obojena dvema bojama, '
            f'podeljena na dan promene</span></div><div class="pc-list">{pc_items}</div></div>')
    else:
        promene_bar_html = ''

    if mix_lista:
        mix_items = ''
        for mx in mix_lista:
            nv = mx['nivoi']
            boje, poc, poz = [], "0%", 0.0
            for idx, (cena, qty, na, arts) in enumerate(nv):
                col = get_color(cena, mx['grupa']) or '#e8ecf4'
                poz += 100.0 / len(nv)
                kraj = "100%" if idx == len(nv) - 1 else f"{poz:.2f}%"
                boje.append(f"{col} {poc},{col} {kraj}")
                if idx < len(nv) - 1:
                    boje.append(f"#1a1a2e {kraj},#1a1a2e calc({kraj} + 2px)")
                    poc = f"calc({kraj} + 2px)"
            sw = "linear-gradient(90deg," + ",".join(boje) + ")"
            cene_txt = ''
            for (cena, qty, na, arts) in nv:
                col = get_color(cena, mx['grupa']) or '#e8ecf4'
                imena = ", ".join(arts[:5]) + ("…" if len(arts) > 5 else "")
                cene_txt += ('<span class="mx-cena">'
                             f'<span class="mx-chip" style="background:{col};'
                             f'color:{text_color(col)}">{fmtnum(cena)} RSD</span>'
                             f'<span class="mx-n">{na} ukusa · {fmtnum(qty)} kom</span>'
                             f'<span class="mx-art">{esc(imena)}</span></span>')
            mix_items += (f'<div class="mx-item"><div class="mx-head">'
                          f'<span class="pc-swatch" style="background:{sw}"></span>'
                          f'<span class="pc-sys">{esc(mx["sistem"])}</span>'
                          f'<span class="pc-gr">{esc(mx["grupa"])}</span></div>'
                          f'<div class="mx-cene">{cene_txt}</div></div>')
        mix_bar_html = (
            '<div class="pc-bar mx-bar"><div class="pc-bar-title">◧ RAZLIČITE CENE UNUTAR GRUPE — '
            f'{esc(poslednji_naziv)} <span class="pc-bar-sub">({len(mix_lista)}) — ćelija je '
            f'obojena na onoliko delova koliko ima cena, širina dela = udeo prodate količine '
            f'po toj ceni</span></div><div class="pc-list">{mix_items}</div></div>')
    else:
        mix_bar_html = ''

    last3_names = nazivi[-3:]
    mag_info = f"Prosek prodaje: {', '.join(last3_names)}"
    cards_html = f'''<div class="mag-cards">
  <div class="mag-card"><div class="mc-label">Ukupno magacin</div>
    <div class="mc-val" style="color:var(--ac)">{fmtnum(total_mag)}</div>
    <div class="mc-sub">komada</div></div>
  <div class="mag-card"><div class="mc-label">Ø Mesečna prodaja</div>
    <div class="mc-val" style="color:var(--t1)">{fmtnum(total_avg)}</div>
    <div class="mc-sub">kom/mesec</div></div>
  <div class="mag-card"><div class="mc-label">Pokrivenost</div>
    <div class="mc-val" style="color:{'var(--grn)' if total_days > 90 else 'var(--red)'}">
      {total_months_str} mes</div><div class="mc-sub">{total_days} dana</div></div>
  <div class="mag-card"><div class="mc-label">Artikala</div>
    <div class="mc-val" style="color:var(--t2)">{len(mag_data)}</div>
    <div class="mc-sub">u magacinu</div></div>
</div>'''
    grupa_cards = '<div class="mag-cards">'
    for g in sorted(mag_grupe_summary.keys()):
        gs = mag_grupe_summary[g]
        dd = round(gs['kol'] / (gs['avg'] / 30)) if gs['avg'] > 0 else 9999
        mm2 = f"{dd / 30:.1f}" if dd < 9999 else "∞"
        clr = 'var(--grn)' if dd > 90 else ('var(--red)' if dd <= 30 else '#92400e')
        grupa_cards += (f'<div class="mag-card"><div class="mc-label">{esc(g)}</div>'
                        f'<div class="mc-val" style="color:{clr}">{mm2} mes</div>'
                        f'<div class="mc-sub">{fmtnum(gs["kol"])} kom · ø {fmtnum(gs["avg"])}/mes'
                        f'</div></div>')
    grupa_cards += '</div>'

    legend_items = ''
    for gnm in sorted(grupa_palete):
        lv = grupa_palete[gnm]
        if not lv:
            continue
        red = grupa_redovna.get(gnm)
        chips = ''
        for c, col in lv:
            ttl = ' title="redovna cena"' if (red is not None and c == red) else ' title="sniženje"'
            chips += (f'<span class="lc" style="background:{col};color:{text_color(col)}"{ttl}>'
                      f'{c}</span>')
        if red is not None:
            chips += ('<span class="lc lc-white" title="sve iznad redovne cene ostaje belo">&gt; '
                      + str(red) + '</span>')
        legend_items += f'<span class="lg-grupa"><b>{esc(gnm)}:</b>{chips}</span>'
    mx_legend = ('<span class="lg-grupa"><b>Različite cene u grupi:</b>'
                 '<span class="lc" style="background:linear-gradient(90deg,#90EE90 0%,#90EE90 50%,'
                 '#1a1a2e 50%,#1a1a2e calc(50% + 2px),#C71585 calc(50% + 2px),#C71585 100%);'
                 'color:#fff">cena A ▏ cena B</span></span>') if mix_lista else ''
    pc_legend = ('<span class="lg-grupa"><b>Promena u toku meseca:</b>'
                 '<span class="lc" style="background:linear-gradient(90deg,#FFB6C1 0%,#FFB6C1 50%,'
                 '#1a1a2e 50%,#1a1a2e calc(50% + 2px),#C71585 calc(50% + 2px),#C71585 100%);'
                 'color:#fff">stara ▏ nova</span></span>') if promene_lista else ''

    if potpun:
        tabs_block = """<div class="tab active" onclick="showTab('prodaja')">📊 PRODAJA</div>
  <div class="tab" onclick="showTab('drv')">🎯 USPEŠNOST AKCIJE</div>
  <div class="tab" onclick="showTab('profit')">💰 PROFITABILNOST</div>
  <div class="tab" onclick="showTab('zalihe')">📦 ZALIHE</div>"""
        filters_profit_block = (
            '<div class="filters hidden" id="filters-profit"><label>Sistem:</label>'
            f'<select id="fp-sistem" onchange="applyProfitFilters()">'
            f'<option value="">Svi sistemi</option>{sistem_options}</select>'
            f'<span style="width:12px"></span><label>Troškovi:</label>{trosak_checks}'
            '<button class="reset-btn" onclick="resetProfitFilters()">✕ Reset</button></div>')
        filters_zalihe_block = (
            '<div class="filters hidden" id="filters-zalihe"><label>Sistem:</label>'
            f'<select id="fz-sistem" onchange="applyZaliheFilters()">'
            f'<option value="">Svi sistemi</option>{sistem_options}</select>'
            '<button class="reset-btn" onclick="document.getElementById(\'fz-sistem\').value=\'\';'
            'applyZaliheFilters()">✕ Reset</button></div>')
        profit_panel_block = (
            f'<div class="panel" id="panel-profit"><div class="tw"><table><thead>{prh}</thead>'
            f'<tbody id="tbody-profit">{chr(10).join(profit_rows)}</tbody></table></div>'
            '<div class="ft">Klikni na sistem → detalji profita i troškova</div></div>')
        zalihe_panel_block = f'''<div class="panel" id="panel-zalihe">
<div class="mag-section"><h3>STANJE MAGACINA</h3><div class="mag-sub">{mag_info}</div>
{cards_html}{grupa_cards}<div class="tw" style="padding:0"><table class="mag-table">
<thead>{mh}</thead><tbody>{chr(10).join(mag_rows)}</tbody></table></div></div>
<div class="zal-divider">ZALIHE PO SISTEMIMA (poslednje stanje)</div>
<div class="tw"><table><thead>{zsh}</thead><tbody id="tbody-zalihe">
{chr(10).join(zal_rows)}</tbody></table></div>
{obrt_lagera_html}
</div>'''
        filters_drv_block = (
            '<div class="filters hidden" id="filters-drv"><label>Sistem:</label>'
            f'<select id="fd-sistem" onchange="applyDrvFilters()">'
            f'<option value="">Svi sistemi</option>{sistem_options}</select>'
            f'<span style="width:12px"></span><label>Troškovi:</label>{trosak_checks_drv}'
            '<button class="reset-btn" onclick="resetDrvFilters()">✕ Reset</button>'
            '<span class="f-hint-drv">svi sistemi po odjavi · uvek svi meseci</span></div>')
        drv_panel_block = (
            f'<div class="panel" id="panel-drv"><div class="tw"><table><thead>{drvh}</thead>'
            f'<tbody id="tbody-drv">{chr(10).join(drv_rows)}</tbody></table></div>'
            '<div class="ft">Profitabilnost po odjavi · svi troškovi (uključi/isključi gore) · '
            'uvek svi meseci · klikni na sistem za razradu</div>'
            f'{akcija_analiza_html}</div>')
    else:
        tabs_block = """<div class="tab active" onclick="showTab('prodaja')">📊 PRODAJA</div>"""
        filters_profit_block = filters_zalihe_block = ''
        profit_panel_block = zalihe_panel_block = ''
        filters_drv_block = drv_panel_block = ''

    html = f'''<!DOCTYPE html>
<html lang="sr"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Dashboard Akcija Vape</title>
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600;700&family=Outfit:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
<style>{CSS}</style></head><body>
<div class="hdr"><div><h1>VAPE DASHBOARD</h1><div class="sub">{info}</div></div>
  <div class="hb" style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">{badge_html}
    <button onclick="toggleAll(true)" style="background:var(--acd);color:var(--ac);
      border:1px solid rgba(37,99,235,0.15)">Otvori sve</button>
    <button onclick="toggleAll(false)" style="background:rgba(90,95,122,0.06);color:var(--t2);
      border:1px solid var(--bd2)">Zatvori sve</button></div></div>
<div class="toolbar"><div class="tabs">
  {tabs_block}
</div></div>
<div class="filters" id="filters-prodaja"><label>Filteri:</label>{sistem_pop}{grupa_pop}
<label>Godina:</label><select id="f-godina" onchange="yearChanged()">
<option value="">Sve godine</option>{godina_options}</select>
<span class="ms-wrap" id="ms-wrap"><button class="ms-btn" onclick="togglePop(event,'ms-pop')">
Meseci: <b id="mesec-count">{num_s}/{num_s}</b> ▾</button>
<div class="ms-pop hidden" id="ms-pop"><div id="mesec-checks">{mesec_checks}</div>
<div class="ms-actions"><button class="mini-btn" onclick="allMonths(true)">Svi</button>
<button class="mini-btn" onclick="allMonths(false)">Nijedan</button>
<button class="mini-btn ms-close" onclick="togglePop(event,'ms-pop')">Zatvori</button></div></div>
</span><button class="reset-btn" onclick="resetFilters()">✕ Reset</button>
<span class="f-hint-drv">može više sistema i više grupa odjednom</span></div>
{filters_profit_block}
{filters_drv_block}
{filters_zalihe_block}
<div class="legend" id="leg"><span class="lg-note">CENE PO GRUPAMA:</span>
{legend_items}{pc_legend}{mx_legend}</div>
<div id="pc-bar-wrap">{promene_bar_html}{mix_bar_html}</div>
<div class="panel active" id="panel-prodaja"><div class="tw"><table><thead>{ph}</thead>
<tbody id="tbody-prodaja">{chr(10).join(prodaja_rows)}</tbody></table></div>
<div class="ft">Klikni na sistem → grupe · Klikni na grupu → artikli</div>{analitika_html}</div>
{profit_panel_block}
{drv_panel_block}
{zalihe_panel_block}
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<script>{_js(prodaja_json, profit_json, drv_json, num_s, num_p, num_d,
             analitika_data_json, akcija_data_json, obrt_chart_json)}</script>
</body></html>'''

    # ---------- Excel + sažetak za direktorske kartice ----------
    xlsx_bytes = _napravi_xlsx(nazivi, sistemi_lista, prodaja_data_js, grand_totals,
                               profit_data_js, nazivi_profit, trosak_ids_list,
                               trosak_names_list, mag_data, total_mag, total_avg,
                               xls_prodaja=xls_prodaja, potpun=potpun,
                               mesec_lbl=poslednji_naziv,
                               palete=grupa_palete)
    summary = {"nazivi": nazivi, "po_sistemu": {}}
    for sistem in sistemi_lista:
        grupe = prodaja_data_js.get(sistem, {})
        total = [0] * num_s
        for vals in grupe.values():
            for i in range(num_s):
                total[i] += vals[i]
        summary["po_sistemu"][sistem] = {"total": total, "grupe": grupe}
    return (html, xlsx_bytes, poslednji_naziv, summary)


def _hex(c):
    """'#aabbcc' -> 'AABBCC' (openpyxl boja), None ostaje None."""
    if not c:
        return None
    c = str(c).strip().lstrip("#")
    return c.upper()[:6] if len(c) >= 6 else None


def _tekst_na(boja):
    """Crn ili beo tekst, u zavisnosti od svetline pozadine."""
    try:
        r, g, b = int(boja[0:2], 16), int(boja[2:4], 16), int(boja[4:6], 16)
        return "FFFFFF" if (0.299 * r + 0.587 * g + 0.114 * b) < 150 else "1A1A2E"
    except Exception:
        return "1A1A2E"


def _napravi_xlsx(nazivi, sistemi, prodaja_data, grand_totals, profit_data, nazivi_profit,
                  trosak_ids, trosak_names, mag_data, total_mag, total_avg,
                  xls_prodaja=None, potpun=True, mesec_lbl="", palete=None):
    """Excel: list „Prodaja" je ISTA tabela kao na ekranu — sistem ▸ grupa ▸ artikal,
    meseci u kolonama, ćelije obojene po ceni (kao u pregledaču), redovi se skupljaju
    i šire kao pivot. Iznad tabele stoji legenda cena po grupama. U potpunom
    izveštaju tu su i listovi Neto profit i Magacin."""
    from openpyxl import Workbook
    from openpyxl.comments import Comment
    from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    hf = PatternFill("solid", fgColor="2563EB")
    hfn = Font(bold=True, color="FFFFFF", size=10, name="Arial")
    bold = Font(bold=True, size=10, name="Arial")
    _thin = Side(style="thin", color="E2E8F0")
    _bord = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
    _sfill = PatternFill("solid", fgColor="EEF2FF")
    _tfill = PatternFill("solid", fgColor="1E293B")
    _siv = Font(size=9, color="7C8494", name="Arial")
    _num = "#,##0"

    ws = wb.active
    ws.title = "Prodaja"
    _nk = len(nazivi)
    _ukupno_kol = 2 + _nk                      # A = naziv, pa meseci, pa TOTAL
    ws.merge_cells("A1:" + get_column_letter(_ukupno_kol) + "1")
    _t = ws["A1"]
    _t.value = ("Prodaja po sistemima, grupama i artiklima"
                + ((" · poslednji mesec " + str(mesec_lbl)) if mesec_lbl else "")
                + ("" if potpun else "  ·  NEPOTPUN IZVEŠTAJ (troškovi nisu učitani)"))
    _t.font = Font(bold=True, size=13, color="1E3A8A", name="Arial")
    ws.row_dimensions[1].height = 22

    # ---- LEGENDA CENA iznad tabele (kao traka „Cene po grupama" na ekranu) ----
    _red = 2
    if palete:
        _red += 1
        _c = ws.cell(row=_red, column=1, value="CENE PO GRUPAMA  (boja ćelije = cena)")
        _c.font = Font(bold=True, size=10, color="6D28D9", name="Arial")
        for _g in sorted(palete.keys()):
            _niv = [x for x in (palete.get(_g) or []) if x and x[0]]
            if not _niv:
                continue
            _red += 1
            _cg = ws.cell(row=_red, column=1, value=str(_g))
            _cg.font = Font(bold=True, size=10, color="334155", name="Arial")
            _cg.alignment = Alignment(horizontal="right", vertical="center")
            for _j, (_cena, _boja) in enumerate(_niv[:_nk]):
                _cc = ws.cell(row=_red, column=2 + _j, value=int(_cena))
                _cc.number_format = _num
                _cc.alignment = Alignment(horizontal="center")
                _bh = _hex(_boja)
                if _bh:
                    _cc.fill = PatternFill("solid", fgColor=_bh)
                    _cc.font = Font(bold=True, size=10, color=_tekst_na(_bh), name="Arial")
                _cc.border = _bord
        _red += 1
        _c = ws.cell(row=_red, column=1,
                     value="Ćelija sa komentarom (crveni ugao) = cena menjana u toku meseca "
                           "ili ukusi nisu svi na istoj ceni — objašnjenje je u komentaru.")
        _c.font = Font(italic=True, size=9, color="7C8494", name="Arial")

    _red += 1                                   # prazan red između legende i tabele
    _hr = _red + 1                              # red zaglavlja tabele
    for _j, _v in enumerate(["SISTEM / GRUPA / ARTIKAL"] + list(nazivi) + ["TOTAL"]):
        _c = ws.cell(row=_hr, column=1 + _j, value=_v)
        _c.font = hfn
        _c.fill = hf
        _c.alignment = Alignment(horizontal="center", wrap_text=True)
        _c.border = _bord
    ws.cell(row=_hr, column=1).alignment = Alignment(horizontal="left", vertical="center")

    def _upis(red, vrednosti, total):
        for j, v in enumerate(vrednosti):
            _c = ws.cell(row=red, column=2 + j, value=int(v or 0))
            _c.number_format = _num
            _c.alignment = Alignment(horizontal="right")
            _c.border = _bord
        _c = ws.cell(row=red, column=2 + len(vrednosti), value=int(total or 0))
        _c.number_format = _num
        _c.alignment = Alignment(horizontal="right")
        _c.border = _bord
        return _c

    if xls_prodaja:
        for _s in xls_prodaja:
            ws.append([str(_s["sistem"])])
            _r = ws.max_row
            _a = ws.cell(row=_r, column=1)
            _a.font = Font(bold=True, size=11, color="1E293B", name="Arial")
            _a.fill = _sfill
            _a.border = _bord
            _tc = _upis(_r, _s["vals"], sum(_s["vals"]))
            for _ci in range(2, _ukupno_kol + 1):
                _c = ws.cell(row=_r, column=_ci)
                _c.fill = _sfill
                _c.font = Font(bold=True, size=10, name="Arial")
            _tc.font = Font(bold=True, size=10, color="1E3A8A", name="Arial")
            for _g in _s["grupe"]:
                ws.append(["    " + str(_g["grupa"])])
                _rg = ws.max_row
                ws.row_dimensions[_rg].outlineLevel = 1
                ws.row_dimensions[_rg].hidden = True
                _ag = ws.cell(row=_rg, column=1)
                _ag.font = Font(bold=True, size=10, color="334155", name="Arial")
                _ag.border = _bord
                _upis(_rg, _g["vals"], sum(_g["vals"]))
                for j, _b in enumerate(_g.get("boje") or []):
                    _bh = _hex(_b)
                    _cc = ws.cell(row=_rg, column=2 + j)
                    if _bh:
                        _cc.fill = PatternFill("solid", fgColor=_bh)
                        _cc.font = Font(bold=True, size=10, color=_tekst_na(_bh), name="Arial")
                    _nt = (_g.get("note") or [""] * len(_g["vals"]))[j] if j < len(
                        _g.get("note") or []) else ""
                    if _nt:
                        try:
                            _cm = Comment(_nt, "Izveštaj prodaje")
                            _cm.width, _cm.height = 260, 90
                            _cc.comment = _cm
                        except Exception:
                            pass
                for _art in _g.get("artikli") or []:
                    ws.append(["        " + str(_art["naziv"])])
                    _ra = ws.max_row
                    ws.row_dimensions[_ra].outlineLevel = 2
                    ws.row_dimensions[_ra].hidden = True
                    ws.cell(row=_ra, column=1).font = _siv
                    ws.cell(row=_ra, column=1).border = _bord
                    _upis(_ra, _art["vals"], sum(_art["vals"]))
                    for j in range(len(_art["vals"])):
                        ws.cell(row=_ra, column=2 + j).font = _siv
                    ws.cell(row=_ra, column=2 + len(_art["vals"])).font = _siv
        ws.append(["TOTAL"])
        _rt = ws.max_row
        _upis(_rt, list(grand_totals), sum(grand_totals))
        for _ci in range(1, _ukupno_kol + 1):
            _c = ws.cell(row=_rt, column=_ci)
            _c.fill = _tfill
            _c.font = Font(bold=True, size=11, color="FFFFFF", name="Arial")
            _c.border = _bord
        ws.sheet_properties.outlinePr.summaryBelow = False
    else:
        ws.append(["Sistem", "Grupa"] + list(nazivi) + ["TOTAL"])
        for sistem in sistemi:
            grupe = prodaja_data.get(sistem, {})
            for g, vals in sorted(grupe.items()):
                ws.append([sistem, g] + list(vals) + [sum(vals)])
    ws.freeze_panes = "B" + str(_hr + 1)
    ws.column_dimensions["A"].width = 46
    for i in range(2, 2 + _nk + 1):
        ws.column_dimensions[get_column_letter(i)].width = 11

    # List „Neto profit" postoji SAMO u potpunom izveštaju — u nepotpunom troškovi
    # nisu učitani, pa bi brojke bile pogrešne.
    if potpun:
        ws2 = wb.create_sheet("Neto profit")
        ws2.append(["Sistem"] + list(nazivi_profit) + ["TOTAL"])
        for c in ws2[1]:
            c.font = hfn
            c.fill = hf
            c.alignment = Alignment(horizontal="center", wrap_text=True)
        for sistem in sistemi:
            pd_s = profit_data.get(sistem) or {}
            n = len(nazivi_profit)
            uk = list(pd_s.get("mkt", [0] * n))
            for tid in trosak_ids[1:]:
                src = pd_s.get("dod" if tid == "dod" else tid, [0] * n)
                for j in range(n):
                    uk[j] += src[j]
            prof = pd_s.get("profit", [0] * n)
            neto = [prof[j] - uk[j] for j in range(n)]
            ws2.append([sistem] + neto + [sum(neto)])
        ws2.column_dimensions["A"].width = 26
        for i in range(2, 2 + len(nazivi_profit) + 1):
            ws2.column_dimensions[get_column_letter(i)].width = 13

    if mag_data and potpun:
        ws3 = wb.create_sheet("Magacin")
        ws3.append(["Artikal", "Magacin (kom)", "Ø prodaja/mes", "Dana zaliha"])
        for c in ws3[1]:
            c.font = hfn
            c.fill = hf
        for art_name, kol, avg_sale, days in mag_data:
            ws3.append([art_name, kol, avg_sale, (days if days < 9999 else "")])
        ws3.append(["UKUPNO", total_mag, total_avg, ""])
        for c in ws3[ws3.max_row]:
            c.font = bold
        ws3.column_dimensions["A"].width = 52
        for col in "BCD":
            ws3.column_dimensions[col].width = 16

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


CSS = '''*{margin:0;padding:0;box-sizing:border-box}
:root{--bg:#f4f6fb;--bg2:#ffffff;--bd:#e2e6ef;--bd2:#d0d5e0;--t1:#1a1a2e;--t2:#5a5f7a;--t3:#8b90a5;--ac:#2563eb;--acd:rgba(37,99,235,0.08);--red:#dc2626;--redd:rgba(220,38,38,0.06);--grn:#16a34a;--grnd:rgba(22,163,74,0.06);--shadow:0 1px 3px rgba(0,0,0,0.06)}
body{font-family:'Outfit',sans-serif;background:var(--bg);color:var(--t1);min-height:100vh}
.hdr{padding:20px 28px 16px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:12px;background:var(--bg2);border-bottom:1px solid var(--bd);box-shadow:var(--shadow)}
.hdr h1{font-family:'IBM Plex Mono',monospace;font-size:20px;font-weight:700;color:var(--ac);letter-spacing:-0.5px}
.hdr .sub{font-size:12px;color:var(--t3);margin-top:2px}
.badge{padding:4px 10px;border-radius:20px;font-size:10px;font-weight:600;font-family:'IBM Plex Mono',monospace}
.bg{background:var(--grnd);color:var(--grn);border:1px solid rgba(22,163,74,0.2)}
.br{background:var(--redd);color:var(--red);border:1px solid rgba(220,38,38,0.2)}
.toolbar{display:flex;gap:0;padding:0;background:var(--bg2);border-bottom:1px solid var(--bd);align-items:stretch}
.tabs{display:flex;gap:0;padding:0 24px;flex:1}
.tab{padding:14px 24px;font-family:'IBM Plex Mono',monospace;font-size:11px;font-weight:600;cursor:pointer;border-bottom:3px solid transparent;color:var(--t3);transition:all .2s;margin-bottom:-1px;user-select:none}
.tab:hover{color:var(--t1)}.tab.active{color:var(--ac);border-bottom-color:var(--ac)}
.filters{display:flex;gap:12px;align-items:center;flex-wrap:wrap;padding:12px 28px;background:var(--bg);border-bottom:1px solid var(--bd)}
.filters label{font-family:'IBM Plex Mono',monospace;font-size:10px;color:var(--t3);font-weight:600;text-transform:uppercase;letter-spacing:.5px}
.filters select{background:var(--bg2);color:var(--t1);border:1px solid var(--bd2);border-radius:8px;padding:6px 12px;font-family:'Outfit',sans-serif;font-size:12px;cursor:pointer;min-width:150px;box-shadow:var(--shadow)}
.filters select:focus{outline:none;border-color:var(--ac)}
.reset-btn{background:var(--redd);color:var(--red);border:1px solid rgba(220,38,38,0.15);padding:6px 14px;border-radius:8px;cursor:pointer;font-size:11px;font-family:'IBM Plex Mono',monospace;font-weight:600}
.mini-btn{background:var(--acd);color:var(--ac);border:1px solid rgba(37,99,235,0.15);padding:5px 11px;border-radius:7px;cursor:pointer;font-size:10px;font-family:'IBM Plex Mono',monospace;font-weight:600}
.mini-btn:hover{background:var(--ac);color:#fff}
.tcb,.mcb{display:inline-flex;align-items:center;gap:4px;font-family:'IBM Plex Mono',monospace;font-size:10px;color:var(--t2);cursor:pointer;padding:3px 8px;border-radius:6px;background:var(--bg2);border:1px solid var(--bd);transition:all .15s;text-transform:none;letter-spacing:0}
.tcb:hover,.mcb:hover{border-color:var(--ac)}.tcb input,.mcb input{accent-color:var(--ac);cursor:pointer}
.tcb input:not(:checked)+span,.mcb input:not(:checked)+span{color:var(--t3);text-decoration:line-through}
.mcb input:checked+span{color:var(--t1);font-weight:600}
.ms-wrap{position:relative;display:inline-block}
.ms-btn{background:var(--bg2);color:var(--t1);border:1px solid var(--bd2);border-radius:8px;padding:6px 12px;cursor:pointer;font-family:'IBM Plex Mono',monospace;font-size:11px;font-weight:600;box-shadow:var(--shadow);white-space:nowrap}
.ms-btn:hover{border-color:var(--ac);color:var(--ac)}
.ms-btn b{color:var(--ac)}
.ms-pop{position:absolute;top:calc(100% + 6px);left:0;z-index:80;background:var(--bg2);border:1px solid var(--bd2);border-radius:10px;box-shadow:0 8px 26px rgba(0,0,0,.14);padding:12px 14px;min-width:330px;max-height:330px;overflow:auto}
.ms-god{margin-bottom:10px}
.ms-god-hdr{display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;gap:8px}
.ms-god-nm{font-family:'IBM Plex Mono',monospace;font-size:11px;font-weight:700;color:var(--ac)}
.ms-god-hdr .mini-btn{padding:2px 7px;font-size:9px;margin-left:4px}
.ms-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:4px}
.ms-grid2{display:grid;grid-template-columns:repeat(2,1fr);gap:4px;margin-bottom:9px}
.ms-pop-wide{min-width:440px}
.ms-grid2 .mcb{justify-content:flex-start;padding:3px 7px;overflow:hidden}
.ms-grid2 .mcb span{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.ms-grid .mcb{justify-content:flex-start;padding:3px 6px}
.ms-actions{display:flex;gap:6px;border-top:1px solid var(--bd);padding-top:9px;margin-top:2px}
.ms-close{margin-left:auto;background:rgba(90,95,122,0.07);color:var(--t2);border-color:var(--bd2)}
.legend{display:flex;gap:8px;align-items:center;flex-wrap:wrap;padding:8px 28px;font-size:10px;color:var(--t3);background:var(--bg);border-bottom:1px solid var(--bd)}
.legend b{font-family:'IBM Plex Mono',monospace;font-weight:700}
.lg-grupa{display:inline-flex;align-items:center;gap:3px;margin-right:14px;padding:2px 0}
.lg-grupa b{color:var(--t2);margin-right:3px}
.lg-note{font-family:'IBM Plex Mono',monospace;font-weight:700;color:var(--ac);letter-spacing:.5px;margin-right:6px}
.lc-white{background:#fff;color:var(--t3);border:1px dashed var(--bd2);box-shadow:none}
.f-hint-drv{font-family:'IBM Plex Mono',monospace;font-size:10px;color:var(--t3);font-style:italic;margin-left:6px}
.lc{display:inline-block;padding:2px 7px;border-radius:4px;font-weight:700;font-size:9px;font-family:'IBM Plex Mono',monospace;color:#1a1a2e;box-shadow:0 1px 2px rgba(0,0,0,0.1)}
.panel{display:none}.panel.active{display:block}
.tw{overflow-x:auto;padding:4px 0 10px;margin:0 16px 26px}
.mag-section .tw{margin-left:0;margin-right:0}
.tw table{overflow:visible}
.tw table thead th:first-child{border-top-left-radius:12px}
.tw table thead th:last-child{border-top-right-radius:12px}
.tw table td:first-child,.tw table th:first-child{position:sticky;left:0;z-index:6;border-right:2px solid var(--bd2)}
.tw table thead th:first-child{top:0;z-index:30;background:#f1f4fb}
.tw table tr.sr td:first-child{background:#eaf1fd}
.tw table tr.gr td:first-child{background:#f3f7fe}
.tw table tr.ar td:first-child{background:#ffffff}
.tw table tr.cr td:first-child{background:#fdf3f3}
.tw table tr.ctr td:first-child{background:#fbeced}
.tw table tr.nr td:first-child{background:#edf8f1}
.tw table tr.mag-row td:first-child{background:#ffffff}
.tw table tr.totalrow td:first-child{background:#e9f0fa}
.tw table tr.sep td{position:static;border-right:none}
.tw table tr.sr:hover td:first-child{background:#dfe9fb}
.tw table tr.gr:hover td:first-child{background:#e9f0fd}
.tw table tr.ar:hover td:first-child,.tw table tr.mag-row:hover td:first-child{background:#f4f5f8}
.tw-top{overflow-x:auto;overflow-y:hidden;margin:2px 16px 0;height:16px}
.tw-top-inner{height:1px}
.tw::-webkit-scrollbar,.tw-top::-webkit-scrollbar{height:15px}
.tw::-webkit-scrollbar-track,.tw-top::-webkit-scrollbar-track{background:#e6eaf4;border-radius:9px}
.tw::-webkit-scrollbar-thumb,.tw-top::-webkit-scrollbar-thumb{background:#8fa2c6;border-radius:9px;border:3px solid #e6eaf4}
.tw::-webkit-scrollbar-thumb:hover,.tw-top::-webkit-scrollbar-thumb:hover{background:var(--ac)}
.tw,.tw-top{scrollbar-width:auto;scrollbar-color:#8fa2c6 #e6eaf4}
table{border-collapse:collapse;width:100%;font-size:11px;margin-top:4px;background:var(--bg2);border-radius:12px;overflow:hidden;box-shadow:var(--shadow)}
thead th{background:linear-gradient(180deg,#f8f9fd,#eef0f7);color:var(--t2);font-weight:700;font-size:10px;font-family:'IBM Plex Mono',monospace;text-transform:uppercase;letter-spacing:.5px;padding:10px 7px;text-align:center;border-bottom:2px solid var(--bd);position:sticky;top:0;z-index:20;white-space:nowrap}
td{padding:5px 7px;border-bottom:1px solid var(--bd);white-space:nowrap;font-family:'IBM Plex Mono',monospace;font-size:11px}
tr.sr{background:rgba(37,99,235,0.03)}tr.sr:hover{background:rgba(37,99,235,0.07)}tr.sr td{border-bottom:1px solid var(--bd2)}
tr.gr{background:rgba(37,99,235,0.015)}tr.gr:hover{background:rgba(37,99,235,0.05)}
tr.ar{background:var(--bg2)}tr.ar:hover{background:rgba(0,0,0,0.015)}
tr.cr{background:rgba(220,38,38,0.02)}tr.ctr{background:rgba(220,38,38,0.04)}tr.nr{background:rgba(22,163,74,0.05)}
tr.sep td{height:4px;background:var(--bg);border:none;padding:0}
.be{display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;border-radius:6px;background:var(--acd);color:var(--ac);font-size:13px;font-weight:800;cursor:pointer;border:none;margin-right:8px;font-family:'IBM Plex Mono',monospace;vertical-align:middle;transition:all .15s}
.be:hover{background:var(--ac);color:#fff}.beg{background:rgba(90,95,122,0.06);color:var(--t2)}.beg:hover{background:var(--t2);color:#fff}
.sn{font-weight:700;color:var(--ac);font-size:12px;cursor:pointer}.gn{color:var(--t2);font-weight:600;cursor:pointer}
.an{padding-left:36px;color:var(--t3);font-size:10px}.an::before{content:'\\21b3 ';color:var(--bd2)}
.n{text-align:right}.nb{text-align:right;font-weight:600}.nt{text-align:right;font-weight:800;color:var(--ac)}
.pc{color:var(--grn)}.pf{color:var(--red)}.cc{color:var(--red);font-style:italic;opacity:.75}
.cl{color:var(--red);font-style:italic;font-size:10px;opacity:.75}
.ctl{color:#b91c1c;font-weight:800;font-size:10px}.ctc{color:#b91c1c;font-weight:700}
.nl{color:var(--grn);font-weight:800;font-size:10px}.np{color:var(--grn);font-weight:700}.nn{color:var(--red);font-weight:700}
.nf{color:var(--red);font-weight:600;font-size:10px}.no{color:var(--ac);font-weight:600;font-size:10px}
.hidden{display:none}
.pc-cell{cursor:help}
.pc-badge{display:inline-block;margin-left:6px;font-size:10px;color:var(--ac);background:var(--acd);border-radius:4px;padding:0 5px;font-family:'IBM Plex Mono',monospace;font-weight:800;cursor:help}
.mix-badge{display:inline-block;margin-left:5px;font-size:11px;color:#92400e;background:rgba(245,158,11,.12);border-radius:4px;padding:0 5px;font-family:'IBM Plex Mono',monospace;font-weight:800;cursor:help}
.mix-cell{cursor:help}
.mx-bar{border-left-color:#f59e0b}
.mx-bar .pc-bar-title{color:#92400e}
.mx-item{font-family:'IBM Plex Mono',monospace;font-size:11px;padding:7px 0;border-bottom:1px dashed var(--bd)}
.mx-item:last-child{border-bottom:none}
.mx-head{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:4px}
.mx-cene{display:flex;flex-direction:column;gap:3px;margin-left:44px}
.mx-cena{display:flex;align-items:baseline;gap:8px;flex-wrap:wrap}
.mx-chip{display:inline-block;min-width:62px;text-align:center;padding:1px 7px;border-radius:4px;font-weight:700;font-size:10px;box-shadow:0 1px 2px rgba(0,0,0,.1)}
.mx-n{color:var(--t2);font-weight:600;white-space:nowrap}
.mx-art{color:var(--t3);font-size:10px}
.pc-bar{margin:8px 16px 0;background:var(--bg2);border-radius:12px;padding:12px 18px;box-shadow:var(--shadow);border-left:4px solid var(--ac)}
.pc-bar-title{font-family:'IBM Plex Mono',monospace;font-size:11px;font-weight:700;color:var(--ac);margin-bottom:8px}
.pc-bar-sub{font-weight:400;color:var(--t3);font-size:10px}
.pc-list{display:flex;flex-direction:column;gap:5px}
.pc-item{display:flex;align-items:center;gap:10px;flex-wrap:wrap;font-family:'IBM Plex Mono',monospace;font-size:11px;padding:4px 0;border-bottom:1px dashed var(--bd)}
.pc-item:last-child{border-bottom:none}
.pc-swatch{display:inline-block;width:34px;height:14px;border-radius:3px;box-shadow:0 1px 2px rgba(0,0,0,.12)}
.pc-sys{font-weight:700;color:var(--t1)}
.pc-gr{color:var(--t2)}
.pc-per{color:var(--ac);font-weight:600}
.pc-dan{color:var(--t3)}
.pc-cene{font-weight:700;margin-left:auto}
.ft{text-align:center;padding:16px;color:var(--t3);font-size:10px;font-family:'IBM Plex Mono',monospace}
.hb button{padding:6px 14px;border-radius:8px;cursor:pointer;font-size:10px;font-family:'IBM Plex Mono',monospace;font-weight:600;transition:all .2s}
tr.totalrow{background:linear-gradient(90deg,rgba(37,99,235,0.06),rgba(22,163,74,0.06))}
tr.totalrow td{border-top:2px solid var(--bd2);padding:9px 7px}
.total-label{font-weight:800;color:var(--ac);font-size:13px;font-family:'IBM Plex Mono',monospace;letter-spacing:1px}
.total-cell{font-size:12px}tr.excluded td{opacity:0.3}
.mag-section{padding:16px 20px;background:var(--bg2);border-radius:12px;margin:8px 16px;box-shadow:var(--shadow)}
.mag-section h3{font-family:'IBM Plex Mono',monospace;font-size:13px;font-weight:700;color:var(--ac);margin-bottom:4px}
.mag-section .mag-sub{font-size:10px;color:var(--t3);margin-bottom:12px;font-family:'IBM Plex Mono',monospace}
.mag-cards{display:flex;gap:12px;margin-bottom:16px;flex-wrap:wrap}
.mag-card{flex:1;min-width:140px;padding:12px 16px;border-radius:10px;background:var(--bg);border:1px solid var(--bd)}
.mag-card .mc-label{font-size:9px;text-transform:uppercase;letter-spacing:.5px;color:var(--t3);font-family:'IBM Plex Mono',monospace;font-weight:600}
.mag-card .mc-val{font-size:20px;font-weight:800;font-family:'IBM Plex Mono',monospace;margin-top:4px}
.mag-card .mc-sub{font-size:9px;color:var(--t3);font-family:'IBM Plex Mono',monospace;margin-top:2px}
table.mag-table{margin-top:0}tr.mag-row{background:var(--bg2)}tr.mag-row:hover{background:rgba(37,99,235,0.03)}
.zal-divider{padding:20px 20px 8px;font-family:'IBM Plex Mono',monospace;font-size:12px;font-weight:700;color:var(--t2);border-bottom:1px solid var(--bd);margin:0 16px}
.analitika-wrap{padding:8px 16px 40px}
.analitika-headline{background:var(--bg2);border-radius:12px;padding:22px 28px;margin-bottom:14px;box-shadow:var(--shadow);border-left:4px solid var(--ac)}
.ah-top{display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;flex-wrap:wrap;gap:8px}
.ah-top h2{font-family:'IBM Plex Mono',monospace;font-size:13px;font-weight:700;color:var(--ac);letter-spacing:1px;margin:0}
.ah-badge{font-family:'IBM Plex Mono',monospace;font-size:9px;color:var(--t3);padding:3px 10px;background:var(--bg);border-radius:6px}
.ah-text{font-size:14px;line-height:1.7;color:var(--t1);margin:0}
.mini-cards-wrap{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:14px}
.mini-card{background:var(--bg2);border-radius:12px;padding:16px 18px;box-shadow:var(--shadow)}
.mc-lbl{font-size:9px;color:var(--t3);font-family:'IBM Plex Mono',monospace;font-weight:700;text-transform:uppercase;letter-spacing:0.8px;margin-bottom:6px}
.mc-row{display:flex;align-items:baseline;gap:8px;margin-bottom:8px}
.mc-pct{font-size:24px;font-weight:800;font-family:'IBM Plex Mono',monospace}
.mc-diff{font-size:11px;color:var(--t3);font-family:'IBM Plex Mono',monospace}
.mc-chart{position:relative;height:60px}
.mc-sub{display:flex;justify-content:space-between;font-size:9px;color:var(--t3);font-family:'IBM Plex Mono',monospace;margin-top:4px}
.big-chart-card{background:var(--bg2);border-radius:12px;padding:18px 22px;margin-bottom:14px;box-shadow:var(--shadow)}
.bc-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;flex-wrap:wrap;gap:8px}
.bc-header h3{font-family:'IBM Plex Mono',monospace;font-size:12px;font-weight:700;color:var(--ac);margin:0}
.bc-legend{display:flex;gap:14px;font-size:10px;font-family:'IBM Plex Mono',monospace;color:var(--t2);flex-wrap:wrap}
.bc-legend span{display:flex;align-items:center;gap:5px}
.lg-line{width:18px;height:3px;background:var(--ac);border-radius:2px}
.lg-dash{width:18px;height:0;border-top:1.5px dashed var(--red)}
.lg-dot{width:8px;height:8px;background:var(--grn);border-radius:50%}
.bc-chart{position:relative;height:260px}
.grupe-section-header{padding:14px 22px;background:var(--bg2);border-radius:12px;box-shadow:var(--shadow);border-left:4px solid var(--t2);margin-bottom:14px}
.grupe-section-header h2{font-family:'IBM Plex Mono',monospace;font-size:13px;font-weight:700;color:var(--t1);letter-spacing:1px;margin:0}
.gsh-sub{font-size:11px;color:var(--t3);font-family:'IBM Plex Mono',monospace;margin-top:2px}
.grupa-block{background:var(--bg2);border-radius:12px;padding:20px 24px;margin-bottom:14px;box-shadow:var(--shadow)}
.grupa-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:14px;flex-wrap:wrap;gap:8px}
.grupa-title-wrap{display:flex;align-items:center;gap:12px;flex-wrap:wrap}
.grupa-pill{padding:5px 12px;border-radius:6px;font-family:'IBM Plex Mono',monospace;font-size:12px;font-weight:700;letter-spacing:0.5px}
.grupa-udeo{font-size:11px;color:var(--t3);font-family:'IBM Plex Mono',monospace}
.grupa-kom{font-family:'IBM Plex Mono',monospace;font-size:18px;font-weight:800}
.grupa-text{font-size:13px;line-height:1.7;color:var(--t1);margin:0 0 14px 0}
.grupa-chart-wrap{position:relative;height:180px;margin-top:8px}
.tb-card,.alarm-card{background:var(--bg2);border-radius:12px;padding:18px 22px;margin-bottom:14px;box-shadow:var(--shadow)}
.tb-header,.alarm-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;flex-wrap:wrap;gap:8px}
.tb-header h3,.alarm-header h3{font-family:'IBM Plex Mono',monospace;font-size:12px;font-weight:700;color:var(--ac);margin:0}
.tb-sub{font-family:'IBM Plex Mono',monospace;font-size:9px;color:var(--t3)}
.tb-grid{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.tb-col{padding:12px 14px;border-radius:8px}
.tb-col-up{background:rgba(22,163,74,0.05);border-left:3px solid var(--grn)}
.tb-col-down{background:rgba(220,38,38,0.05);border-left:3px solid var(--red)}
.tb-col-title{font-size:10px;font-family:'IBM Plex Mono',monospace;font-weight:700;margin-bottom:8px;letter-spacing:0.5px}
.tb-row{display:flex;justify-content:space-between;align-items:center;padding:6px 0;border-bottom:1px solid rgba(0,0,0,0.05);font-size:11px;font-family:'IBM Plex Mono',monospace}
.tb-row:last-child{border-bottom:none}
.tb-name{font-weight:700;color:var(--t1)}
.tb-vals{font-size:11px}
.tb-diff{color:var(--t3);font-size:10px}
.tb-empty{font-size:11px;color:var(--t3);font-style:italic;padding:6px 0;font-family:'IBM Plex Mono',monospace}
.alarm-counts{display:flex;gap:10px;font-family:'IBM Plex Mono',monospace;font-size:10px;font-weight:600;flex-wrap:wrap}
.alarm-counts span{padding:2px 8px;border-radius:6px}
.ac-kritic{background:rgba(220,38,38,0.08);color:var(--red)}
.ac-upoz{background:rgba(245,158,11,0.08);color:#92400e}
.ac-rekord{background:rgba(22,163,74,0.08);color:var(--grn)}
.alarm-list{display:flex;flex-direction:column;gap:8px}
.alarm-row{padding:10px 14px;border-radius:6px;display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap}
.alarm-text{font-size:12px;flex:1;min-width:200px}
.alarm-level{font-size:10px;font-family:'IBM Plex Mono',monospace;font-weight:700;white-space:nowrap}
.alarm-empty{font-size:12px;color:var(--t3);font-style:italic;padding:12px;text-align:center;font-family:'IBM Plex Mono',monospace}
.akcija-analiza-wrap{padding:8px 16px 40px}
.aa-header{background:var(--bg2);border-radius:12px;padding:18px 24px;margin-bottom:14px;box-shadow:var(--shadow);border-left:4px solid var(--grn)}
.aa-header h2{font-family:'IBM Plex Mono',monospace;font-size:13px;font-weight:700;color:var(--grn);letter-spacing:1px;margin:0}
.aa-sub{font-size:11px;color:var(--t3);font-family:'IBM Plex Mono',monospace;margin-top:4px}
.ak-block{background:var(--bg2);border-radius:12px;padding:22px 26px;margin-bottom:14px;box-shadow:var(--shadow)}
.ak-block-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:16px;padding-bottom:14px;border-bottom:1px solid var(--bd);flex-wrap:wrap;gap:8px}
.ak-sistem-info{display:flex;align-items:center;gap:12px;flex-wrap:wrap}
.ak-pill{background:var(--t1);color:#fff;padding:6px 14px;border-radius:6px;font-family:'IBM Plex Mono',monospace;font-size:13px;font-weight:700;letter-spacing:0.5px}
.ak-meta{font-family:'IBM Plex Mono',monospace;font-size:11px;color:var(--t3)}
.ak-status{padding:4px 12px;border-radius:6px;font-family:'IBM Plex Mono',monospace;font-size:12px;font-weight:700;white-space:nowrap}
.ak-zakljucak{border-radius:8px;padding:14px 18px;margin-bottom:18px}
.ak-zakljucak-label{font-family:'IBM Plex Mono',monospace;font-size:9px;font-weight:700;letter-spacing:0.8px;margin-bottom:6px}
.ak-chart-wrap{margin-top:8px}
.ak-chart-legend{display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;flex-wrap:wrap;gap:8px}
.ak-chart-title{font-family:'IBM Plex Mono',monospace;font-size:10px;color:var(--t2);font-weight:700}
.ak-legend-items{display:flex;gap:10px;font-size:9px;font-family:'IBM Plex Mono',monospace;color:var(--t2)}
.ak-legend-items span{display:flex;align-items:center;gap:4px}
.ak-leg-sq{width:10px;height:10px;border-radius:2px}
.ak-leg-akcija{background:#0d9488}
.ak-leg-redovno{background:#cbd5e1}
.ak-leg-profit{display:inline-block;width:18px;height:2px;background:#1e3a8a;border-radius:1px;vertical-align:middle}
.ak-chart{position:relative;height:200px}
.ak-grupe-table-wrap{margin-bottom:18px}
.ak-table-title{font-family:'IBM Plex Mono',monospace;font-size:11px;font-weight:700;color:var(--t2);margin-bottom:8px}
.ak-grupe-table{width:100%;border-collapse:collapse;font-family:'IBM Plex Mono',monospace;font-size:11px;background:var(--bg);border-radius:8px;overflow:hidden}
.ak-grupe-table thead th{background:linear-gradient(180deg,#f8f9fd,#eef0f7);color:var(--t2);font-weight:700;font-size:9px;text-transform:uppercase;letter-spacing:0.5px;padding:7px 8px;border-bottom:1px solid var(--bd);text-align:center}
.ak-grupe-table th.th-cena{color:var(--ac);border-right:1px solid var(--bd)}
.ak-grupe-table th.th-qty{color:var(--grn)}
.ak-grupe-table th.th-akc{background:rgba(22,163,74,0.04)}
.ak-grupe-table th.th-neakc{background:rgba(139,144,165,0.04)}
.ak-grupe-table th.th-raz{background:rgba(37,99,235,0.04)}
.ak-grupe-table tbody td{padding:7px 8px;border-bottom:1px solid var(--bd);background:var(--bg2)}
.ak-grupe-table tbody tr:last-child td{border-bottom:none}
.ak-grupe-table td.g-naziv{text-align:left;font-weight:600;color:var(--t1)}
.ak-grupe-table td.g-num{text-align:right;font-weight:600}
.ak-grupe-table td.g-num:nth-child(4){border-right:1px solid var(--bd)}
.ak-grupe-table .g-broj{color:var(--t3);font-weight:400;font-size:9px}
.ak-neto-row{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;margin-bottom:18px}
.ak-neto-box{padding:12px 14px;border-radius:8px;text-align:center}
.ak-neto-akc{background:rgba(22,163,74,0.05);border:1px solid rgba(22,163,74,0.15)}
.ak-neto-neakc{background:rgba(139,144,165,0.05);border:1px solid rgba(139,144,165,0.15)}
.ak-neto-razlika{background:rgba(37,99,235,0.05);border:1px solid rgba(37,99,235,0.15)}
.ak-neto-lbl{font-family:'IBM Plex Mono',monospace;font-size:9px;color:var(--t3);font-weight:700;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:6px}
.ak-neto-val{font-family:'IBM Plex Mono',monospace;font-size:17px;font-weight:800;color:var(--t1)}
.obrt-wrap{margin:24px 16px 16px 16px}
.obrt-header{background:var(--bg2);border-radius:12px;padding:14px 22px;margin-bottom:14px;box-shadow:0 1px 3px rgba(0,0,0,0.06);border-left:4px solid var(--ac)}
.obrt-header h3{font-family:'IBM Plex Mono',monospace;font-size:13px;font-weight:700;color:var(--ac);letter-spacing:1px;margin:0}
.obrt-sub{font-family:'IBM Plex Mono',monospace;font-size:11px;color:var(--t3);margin-top:4px}
.obrt-card{background:var(--bg2);border-radius:12px;padding:18px 22px;box-shadow:0 1px 3px rgba(0,0,0,0.06)}
.obrt-legend-row{display:flex;justify-content:space-between;align-items:center;margin-bottom:14px;flex-wrap:wrap;gap:8px}
.obrt-title{font-family:'IBM Plex Mono',monospace;font-size:12px;font-weight:700;color:var(--ac)}
.obrt-legend{display:flex;gap:14px;font-size:10px;font-family:'IBM Plex Mono',monospace;color:var(--t2);flex-wrap:wrap}
.obrt-legend>span{display:flex;align-items:center;gap:5px}
.obrt-leg-sq{display:inline-block;width:14px;height:10px;border-radius:2px}
.obrt-chart{position:relative;height:380px}
.obrt-note{font-size:10px;color:var(--t3);font-family:'IBM Plex Mono',monospace;margin-top:10px;text-align:center;font-style:italic}'''


def _js(prodaja_json, profit_json, drv_json, num_s, num_p, num_d,
        analitika_data_json, akcija_data_json, obrt_chart_json):
    """JavaScript dashboarda (filteri, grafikoni) — isti kao u lokalnoj skripti."""
    return ('var PRODAJA_DATA=' + prodaja_json + ';\nvar PROFIT_DATA=' + profit_json +
            ';\nvar DRV_DATA=' + drv_json +
            ';\nvar NUM_MONTHS=' + str(num_s) +
            ';\nvar NUM_PROFIT_MONTHS=' + str(num_p) +
            ';\nvar NUM_DRV_MONTHS=' + str(num_d) + ''';
function showTab(n){document.querySelectorAll('.panel').forEach(function(p){p.classList.remove('active')});document.querySelectorAll('.tab').forEach(function(t){t.classList.remove('active')});var pnl=document.getElementById('panel-'+n);if(pnl)pnl.classList.add('active');document.querySelectorAll('.tab').forEach(function(t){var oc=t.getAttribute('onclick');if(oc&&oc.indexOf("'"+n+"'")>-1)t.classList.add('active')});function _sd(id,v){var e=document.getElementById(id);if(e)e.style.display=v}_sd('leg',(n==='prodaja')?'flex':'none');_sd('filters-prodaja',(n==='prodaja')?'flex':'none');_sd('pc-bar-wrap',(n==='prodaja')?'block':'none');_sd('filters-profit',(n==='profit')?'flex':'none');_sd('filters-drv',(n==='drv')?'flex':'none');_sd('filters-zalihe',(n==='zalihe')?'flex':'none')}
function tog(id){var btn=document.getElementById('b-'+id);if(!btn)return;var isO=btn.textContent.trim()==='\\u2212';var rows=document.querySelectorAll('tr[data-p="'+id+'"]');if(isO){rows.forEach(function(r){r.classList.add('hidden');var cb=r.querySelector('.beg');if(cb){cb.textContent='+';var cid=cb.id.replace('b-','');document.querySelectorAll('tr[data-p="'+cid+'"]').forEach(function(cr){cr.classList.add('hidden')})}});btn.textContent='+'}else{var uProdaji=false;if(rows.length){var tb=rows[0].parentNode;uProdaji=(tb&&tb.id==='tbody-prodaja')}var sG=uProdaji?selSet('grupa-checks'):null;rows.forEach(function(r){var rg=r.getAttribute('data-grupa');if(sG&&rg&&!sG[rg])return;r.classList.remove('hidden')});btn.textContent='\\u2212'}}
function toggleAll(o){var ap=document.querySelector('.panel.active');if(!ap)return;ap.querySelectorAll('.be').forEach(function(btn){var id=btn.id.replace('b-','');var rows=document.querySelectorAll('tr[data-p="'+id+'"]');rows.forEach(function(r){o?r.classList.remove('hidden'):r.classList.add('hidden')});btn.textContent=o?'\\u2212':'+'})}
function fmtN(v){return v.toLocaleString('sr-RS')}
function parseN(s){var t=String(s).replace(/[^0-9\\-]/g,'');return (t===''||t==='-')?0:parseInt(t,10)}
function selSet(id){var o={};document.querySelectorAll('#'+id+' input').forEach(function(cb){if(cb.checked)o[cb.value]=true});return o}
function cntOf(id){var a=document.querySelectorAll('#'+id+' input');var n=0;a.forEach(function(cb){if(cb.checked)n++});return n+'/'+a.length}
function setAll(id,v){document.querySelectorAll('#'+id+' input').forEach(function(cb){cb.checked=v})}
function allSistemi(v){setAll('sistem-checks',v);applyFilters()}
function allGrupe(v){setAll('grupa-checks',v);applyFilters()}
function getActiveMonths(){var o={};document.querySelectorAll('#mesec-checks input').forEach(function(cb){if(cb.checked)o[parseInt(cb.value,10)]=true});return o}
function yearChanged(){var y=document.getElementById('f-godina').value;document.querySelectorAll('#mesec-checks input').forEach(function(cb){cb.checked=(!y)||(cb.getAttribute('data-god')===y)});applyFilters()}
function allMonths(v){setAll('mesec-checks',v);applyFilters()}
function yearToggle(y,v){document.querySelectorAll('#mesec-checks input[data-god="'+y+'"]').forEach(function(cb){cb.checked=v});applyFilters()}
function togglePop(e,id){if(e){e.stopPropagation();e.preventDefault()}
  var p=document.getElementById(id);if(!p)return;var bilo=p.classList.contains('hidden');
  document.querySelectorAll('.ms-pop').forEach(function(x){x.classList.add('hidden')});
  if(bilo)p.classList.remove('hidden');}
document.addEventListener('click',function(e){
  var t=e.target;while(t&&t!==document){if(t.classList&&t.classList.contains('ms-wrap'))return;t=t.parentNode}
  document.querySelectorAll('.ms-pop').forEach(function(x){x.classList.add('hidden')});});
function applyFilters(){
  var selS=selSet('sistem-checks');
  var selG=selSet('grupa-checks');
  var act=getActiveMonths();
  var tbody=document.getElementById('tbody-prodaja');
  if(!tbody)return;
  tbody.querySelectorAll('.be').forEach(function(btn){btn.textContent='+'});
  tbody.querySelectorAll('tr').forEach(function(r){
    var rS=r.getAttribute('data-sistem');
    var isSr=r.classList.contains('sr');
    var isGr=r.classList.contains('gr');
    var isAr=r.classList.contains('ar');
    var isSep=r.classList.contains('sep');
    var isTotal=r.classList.contains('totalrow');
    if(isTotal){r.classList.remove('hidden');return}
    if(rS&&!selS[rS]){r.classList.add('hidden');return}
    if(isSr){
      var cells=r.querySelectorAll('.mcell');
      var allG=PRODAJA_DATA[rS]||{};
      var sums=new Array(NUM_MONTHS).fill(0);var ima=false;
      for(var g in allG){if(!selG[g])continue;ima=true;
        for(var i=0;i<NUM_MONTHS;i++)sums[i]+=allG[g][i]}
      if(!ima){r.classList.add('hidden');return}
      for(var i=0;i<NUM_MONTHS;i++)cells[i].textContent=fmtN(sums[i]);
      r.classList.remove('hidden');return;
    }
    if(isGr||isAr){r.classList.add('hidden');return}
    if(isSep){
      var aG=PRODAJA_DATA[rS]||{};var ima2=false;
      for(var g2 in aG){if(selG[g2]){ima2=true;break}}
      if(!ima2){r.classList.add('hidden');return}
      r.classList.remove('hidden');return;
    }
  });
  recalcTotals(selS,selG);
  for(var i=0;i<NUM_MONTHS;i++){
    var show=!!act[i];
    document.querySelectorAll('#panel-prodaja .m'+i).forEach(function(el){el.style.display=show?'':'none'});
  }
  tbody.querySelectorAll('tr').forEach(function(r){
    var tot=r.querySelector('.totcell');if(!tot)return;
    var cells=r.querySelectorAll('.mcell');var s=0;
    for(var i=0;i<cells.length;i++){if(act[i])s+=parseN(cells[i].textContent)}
    tot.textContent=fmtN(s);
  });
  function _set(id,txt){var e=document.getElementById(id);if(e)e.textContent=txt}
  _set('cnt-sistem',cntOf('sistem-checks'));
  _set('cnt-grupa',cntOf('grupa-checks'));
  _set('mesec-count',Object.keys(act).length+'/'+NUM_MONTHS);
}
function recalcTotals(selS,selG){
  var t=new Array(NUM_MONTHS).fill(0);
  for(var s in PRODAJA_DATA){if(!selS[s])continue;
    for(var g in PRODAJA_DATA[s]){if(!selG[g])continue;
      var v=PRODAJA_DATA[s][g];for(var i=0;i<NUM_MONTHS;i++)t[i]+=v[i]}}
  var tr=document.getElementById('prodaja-total');if(!tr)return;
  var c=tr.querySelectorAll('.mcell');
  for(var i=0;i<NUM_MONTHS;i++)c[i].textContent=fmtN(t[i]);
}
function resetFilters(){
  setAll('sistem-checks',true);setAll('grupa-checks',true);setAll('mesec-checks',true);
  var fy=document.getElementById('f-godina');if(fy)fy.value='';
  applyFilters();
}
function applyProfitFilters(){var fS=document.getElementById('fp-sistem').value;var tbody=document.getElementById('tbody-profit');var checks=document.querySelectorAll('#filters-profit input[type=checkbox]');var ac={};checks.forEach(function(cb){ac[cb.value]=cb.checked});tbody.querySelectorAll('.be').forEach(function(btn){btn.textContent='+'});tbody.querySelectorAll('tr').forEach(function(r){var rS=r.getAttribute('data-sistem');var isSr=r.classList.contains('sr');var isSep=r.classList.contains('sep');var isTotal=r.classList.contains('totalrow');if(isTotal){r.classList.remove('hidden');return}if(isSr||isSep){if(fS&&rS&&rS!==fS)r.classList.add('hidden');else r.classList.remove('hidden');return}r.classList.add('hidden')});tbody.querySelectorAll('tr[data-trosak]').forEach(function(r){var tid=r.getAttribute('data-trosak');r.classList.toggle('excluded',!ac[tid])});var grandT=new Array(NUM_PROFIT_MONTHS).fill(0);tbody.querySelectorAll('.ukupni-row').forEach(function(ukRow){var sid=ukRow.getAttribute('data-sid');var sistem=ukRow.getAttribute('data-sistem');var pd=PROFIT_DATA[sistem];if(!pd)return;var uk=new Array(NUM_PROFIT_MONTHS).fill(0);if(ac['mkt'])for(var i=0;i<NUM_PROFIT_MONTHS;i++)uk[i]+=pd['mkt'][i];if(ac['dod']&&pd['dod'])for(var i=0;i<NUM_PROFIT_MONTHS;i++)uk[i]+=pd['dod'][i];for(var tid in pd){if(tid==='profit'||tid==='mkt'||tid==='dod')continue;if(ac[tid])for(var i=0;i<NUM_PROFIT_MONTHS;i++)uk[i]+=pd[tid][i]}var ukT=uk.reduce(function(a,b){return a+b},0);var cells=ukRow.querySelectorAll('td');for(var i=3;i<3+NUM_PROFIT_MONTHS;i++)cells[i].textContent=fmtN(uk[i-3]);cells[3+NUM_PROFIT_MONTHS].textContent=fmtN(ukT);var nr=tbody.querySelector('.neto-row[data-sid="'+sid+'"]');if(!nr)return;var p=pd['profit'];var nc=nr.querySelectorAll('td');var nt=0;for(var i=0;i<NUM_PROFIT_MONTHS;i++){var nv=p[i]-uk[i];nt+=nv;nc[3+i].textContent=fmtN(nv);nc[3+i].className='n '+(nv>=0?'np':'nn')}nc[3+NUM_PROFIT_MONTHS].textContent=fmtN(nt);nc[3+NUM_PROFIT_MONTHS].className='n '+(nt>=0?'np':'nn');nc[3+NUM_PROFIT_MONTHS].style.fontSize='13px';var sr=tbody.querySelector('.profit-sr[data-sid="'+sid+'"]');if(sr){var sc=sr.querySelectorAll('td');for(var i=3;i<3+NUM_PROFIT_MONTHS;i++){var nv2=p[i-3]-uk[i-3];sc[i].textContent=fmtN(nv2);sc[i].className='nb '+(nv2>=0?'np':'nn')}var stot=p.reduce(function(a,b){return a+b},0)-ukT;sc[3+NUM_PROFIT_MONTHS].textContent=fmtN(stot);sc[3+NUM_PROFIT_MONTHS].className='nb '+(stot>=0?'np':'nn');sc[3+NUM_PROFIT_MONTHS].style.fontSize='13px'}if(!fS||sistem===fS){for(var i=0;i<NUM_PROFIT_MONTHS;i++)grandT[i]+=p[i]-uk[i]}});var tRow=document.getElementById('profit-total');if(tRow){var tc=tRow.querySelectorAll('td');var ggt=grandT.reduce(function(a,b){return a+b},0);for(var i=3;i<3+NUM_PROFIT_MONTHS;i++){tc[i].textContent=fmtN(grandT[i-3]);tc[i].className='nb total-cell '+(grandT[i-3]>=0?'np':'nn')}tc[3+NUM_PROFIT_MONTHS].textContent=fmtN(ggt);tc[3+NUM_PROFIT_MONTHS].className='nb total-cell '+(ggt>=0?'np':'nn');tc[3+NUM_PROFIT_MONTHS].style.fontSize='13px'}}
function resetProfitFilters(){document.getElementById('fp-sistem').value='';document.querySelectorAll('#filters-profit input[type=checkbox]').forEach(function(cb){cb.checked=true});applyProfitFilters()}
function applyDrvFilters(){
  var fS=document.getElementById('fd-sistem').value;
  var tbody=document.getElementById('tbody-drv');
  if(!tbody)return;
  var ac={};
  document.querySelectorAll('#filters-drv input[type=checkbox]').forEach(function(cb){ac[cb.value]=cb.checked});
  tbody.querySelectorAll('.be').forEach(function(btn){btn.textContent='+'});
  tbody.querySelectorAll('tr').forEach(function(r){
    var rS=r.getAttribute('data-sistem');
    var isSr=r.classList.contains('sr');
    var isSep=r.classList.contains('sep');
    var isTotal=r.classList.contains('totalrow');
    if(isTotal){r.classList.remove('hidden');return}
    if(isSr||isSep){if(fS&&rS&&rS!==fS)r.classList.add('hidden');else r.classList.remove('hidden');return}
    r.classList.add('hidden');
  });
  tbody.querySelectorAll('tr[data-trosak]').forEach(function(r){
    var tid=r.getAttribute('data-trosak');
    r.classList.toggle('excluded',!ac[tid]);
  });
  var grandT=new Array(NUM_DRV_MONTHS).fill(0);
  tbody.querySelectorAll('.ukupni-row-dv').forEach(function(ukRow){
    var sid=ukRow.getAttribute('data-sid');
    var sistem=ukRow.getAttribute('data-sistem');
    var dd=DRV_DATA[sistem];
    if(!dd)return;
    var uk=new Array(NUM_DRV_MONTHS).fill(0);
    if(ac['mkt'])for(var i=0;i<NUM_DRV_MONTHS;i++)uk[i]+=dd['mkt'][i];
    if(ac['dod']&&dd['dod'])for(var i=0;i<NUM_DRV_MONTHS;i++)uk[i]+=dd['dod'][i];
    for(var tid in dd){
      if(tid==='profit'||tid==='mkt'||tid==='dod')continue;
      if(ac[tid])for(var i=0;i<NUM_DRV_MONTHS;i++)uk[i]+=dd[tid][i];
    }
    var ukT=uk.reduce(function(a,b){return a+b},0);
    var cells=ukRow.querySelectorAll('td');
    for(var i=3;i<3+NUM_DRV_MONTHS;i++)cells[i].textContent=fmtN(uk[i-3]);
    cells[3+NUM_DRV_MONTHS].textContent=fmtN(ukT);
    var nr=tbody.querySelector('.neto-row-dv[data-sid="'+sid+'"]');
    if(!nr)return;
    var p=dd['profit'];
    var nc=nr.querySelectorAll('td');
    var nt=0;
    for(var i=0;i<NUM_DRV_MONTHS;i++){var nv=p[i]-uk[i];nt+=nv;nc[3+i].textContent=fmtN(nv);nc[3+i].className='n '+(nv>=0?'np':'nn')}
    nc[3+NUM_DRV_MONTHS].textContent=fmtN(nt);
    nc[3+NUM_DRV_MONTHS].className='n '+(nt>=0?'np':'nn');
    nc[3+NUM_DRV_MONTHS].style.fontSize='13px';
    var sr=tbody.querySelector('.drv-sr[data-sid="'+sid+'"]');
    if(sr){
      var sc=sr.querySelectorAll('td');
      for(var i=3;i<3+NUM_DRV_MONTHS;i++){var nv2=p[i-3]-uk[i-3];sc[i].textContent=fmtN(nv2);sc[i].className='nb '+(nv2>=0?'np':'nn')}
      var stot=p.reduce(function(a,b){return a+b},0)-ukT;
      sc[3+NUM_DRV_MONTHS].textContent=fmtN(stot);
      sc[3+NUM_DRV_MONTHS].className='nb '+(stot>=0?'np':'nn');
      sc[3+NUM_DRV_MONTHS].style.fontSize='13px';
    }
    if(!fS||sistem===fS){for(var i=0;i<NUM_DRV_MONTHS;i++)grandT[i]+=p[i]-uk[i]}
  });
  var tRow=document.getElementById('drv-total');
  if(tRow){
    var tc=tRow.querySelectorAll('td');
    var ggt=grandT.reduce(function(a,b){return a+b},0);
    for(var i=3;i<3+NUM_DRV_MONTHS;i++){tc[i].textContent=fmtN(grandT[i-3]);tc[i].className='nb total-cell '+(grandT[i-3]>=0?'np':'nn')}
    tc[3+NUM_DRV_MONTHS].textContent=fmtN(ggt);
    tc[3+NUM_DRV_MONTHS].className='nb total-cell '+(ggt>=0?'np':'nn');
    tc[3+NUM_DRV_MONTHS].style.fontSize='13px';
  }
}
function resetDrvFilters(){document.getElementById('fd-sistem').value='';document.querySelectorAll('#filters-drv input[type=checkbox]').forEach(function(cb){cb.checked=true});applyDrvFilters()}
function applyZaliheFilters(){var fS=document.getElementById('fz-sistem').value;var tbody=document.getElementById('tbody-zalihe');if(!tbody)return;tbody.querySelectorAll('.be').forEach(function(btn){btn.textContent='+'});tbody.querySelectorAll('tr').forEach(function(r){var rS=r.getAttribute('data-sistem');var isSr=r.classList.contains('sr');var isSep=r.classList.contains('sep');var isTotal=r.classList.contains('totalrow');if(isTotal){r.classList.remove('hidden');return}if(isSr||isSep){if(fS&&rS&&rS!==fS)r.classList.add('hidden');else r.classList.remove('hidden');return}r.classList.add('hidden')})}
var ANALITIKA_DATA=''' + analitika_data_json + ''';
var chartsInitialized=false;
function initAnalitikaCharts(){
  if(chartsInitialized||typeof Chart==='undefined')return;
  chartsInitialized=true;
  var d=ANALITIKA_DATA;
  var monoFont={family:"IBM Plex Mono",size:9};
  var monoFont10={family:"IBM Plex Mono",size:10};
  var commonTooltip={backgroundColor:"#fff",titleColor:"#1a1a2e",bodyColor:"#5a5f7a",borderColor:"#e2e6ef",borderWidth:1,padding:8,titleFont:monoFont10,bodyFont:monoFont10};
  function makeMini(id,obj){
    if(!obj)return;
    var el=document.getElementById(id);
    if(!el)return;
    new Chart(el,{
      type:"bar",
      data:{labels:obj.labels,datasets:[{data:obj.values,backgroundColor:["#cbd5e0",obj.color],borderRadius:4}]},
      options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{enabled:false}},scales:{y:{display:false,beginAtZero:false},x:{display:false}}}
    });
  }
  makeMini("mini1",d.mini1);
  makeMini("mini2",d.mini2);
  makeMini("mini3",d.mini3);
  var bigEl=document.getElementById("bigTrendChart");
  if(bigEl){
    var bg=d.big;
    var n=bg.values.length;
    var rolling6=bg.values.map(function(_,i,a){
      if(i<5)return null;
      var s=0;
      for(var k=i-5;k<=i;k++)s+=a[k];
      return Math.round(s/6);
    });
    var yoyArr=new Array(n).fill(null);
    for(var i=0;i<n;i++){
      var prevI=i-12;
      if(prevI>=0)yoyArr[i]=bg.values[prevI];
    }
    new Chart(bigEl,{
      type:"line",
      data:{
        labels:bg.labels,
        datasets:[
          {label:"Mesečna prodaja",data:bg.values,borderColor:"#2563eb",backgroundColor:"rgba(37,99,235,0.08)",fill:true,tension:0.35,borderWidth:2.5,pointRadius:3.5,pointBackgroundColor:"#2563eb",pointBorderColor:"#fff",pointBorderWidth:1.5},
          {label:"6M prosek",data:rolling6,borderColor:"#dc2626",borderDash:[6,4],borderWidth:1.8,pointRadius:0,fill:false,tension:0.35},
          {label:"Isti mesec prethodne godine",data:yoyArr,borderColor:"#16a34a",borderWidth:0,pointRadius:3,pointBackgroundColor:"#16a34a",pointBorderColor:"#fff",pointBorderWidth:1,showLine:false}
        ]
      },
      options:{
        responsive:true,
        maintainAspectRatio:false,
        plugins:{legend:{display:false},tooltip:Object.assign({mode:"index",intersect:false},commonTooltip)},
        scales:{
          y:{beginAtZero:false,grid:{color:"rgba(0,0,0,0.04)"},ticks:{font:monoFont10,color:"#8b90a5"}},
          x:{grid:{display:false},ticks:{font:monoFont,color:"#8b90a5",maxRotation:45}}
        }
      }
    });
  }
  Object.keys(d.grupe).forEach(function(grKey){
    var g=d.grupe[grKey];
    var el=document.getElementById("chart-"+grKey);
    if(!el)return;
    var labels=g.has_yoy?g.labels:g.labels.slice(0,3);
    var values=g.has_yoy?g.values:g.values.slice(0,3);
    var colors=values.map(function(_,i){return i===0?g.main_color:"#cbd5e0";});
    new Chart(el,{
      type:"bar",
      data:{labels:labels,datasets:[{data:values,backgroundColor:colors,borderRadius:6,borderSkipped:false}]},
      options:{
        indexAxis:"y",
        responsive:true,
        maintainAspectRatio:false,
        plugins:{legend:{display:false},tooltip:commonTooltip},
        scales:{
          x:{beginAtZero:false,grid:{color:"rgba(0,0,0,0.04)"},ticks:{font:monoFont10,color:"#8b90a5",callback:function(v){return v.toLocaleString("sr-RS");}}},
          y:{grid:{display:false},ticks:{font:{family:"IBM Plex Mono",size:10,weight:"bold"},color:"#1a1a2e"}}
        }
      }
    });
  });
}
var AKCIJA_DATA=''' + akcija_data_json + ''';
var OBRT_DATA=''' + obrt_chart_json + ''';
var akcijaChartsInit=false;
function initAkcijaCharts(){
  if(akcijaChartsInit||typeof Chart==='undefined')return;
  akcijaChartsInit=true;
  var monoFont={family:"IBM Plex Mono",size:9};
  var monoFont10={family:"IBM Plex Mono",size:10};
  var commonTooltip={backgroundColor:"#fff",titleColor:"#1a1a2e",bodyColor:"#5a5f7a",borderColor:"#e2e6ef",borderWidth:1,padding:8,titleFont:monoFont10,bodyFont:monoFont10};
  Object.keys(AKCIJA_DATA).forEach(function(safeId){
    var d=AKCIJA_DATA[safeId];
    var el=document.getElementById("ak-chart-"+safeId);
    if(!el)return;
    var colors=d.values.map(function(_,i){return d.is_akcija[i]?"#0d9488":"#cbd5e1";});
    new Chart(el,{
      data:{
        labels:d.labels,
        datasets:[
          {type:"bar",label:"Količina",data:d.values,backgroundColor:colors,borderRadius:3,yAxisID:"y",order:2},
          {type:"line",label:"Bruto profit",data:d.profit,borderColor:"#1e3a8a",backgroundColor:"#1e3a8a",borderWidth:2.5,pointRadius:3,pointBackgroundColor:"#1e3a8a",pointBorderColor:"#fff",pointBorderWidth:1,tension:0.3,fill:false,yAxisID:"y1",order:1}
        ]
      },
      options:{
        responsive:true,maintainAspectRatio:false,
        plugins:{
          legend:{display:false},
          tooltip:Object.assign({mode:"index",intersect:false},commonTooltip,{callbacks:{label:function(c){
            if(c.dataset.type==="line")return "Bruto profit: "+Math.round(c.parsed.y).toLocaleString("sr-RS")+" RSD";
            return (d.is_akcija[c.dataIndex]?"AKCIJA: ":"redovno: ")+c.parsed.y.toLocaleString("sr-RS")+" kom";
          }}})
        },
        scales:{
          y:{position:"left",beginAtZero:true,grid:{color:"rgba(0,0,0,0.04)"},ticks:{font:monoFont,color:"#8b90a5"},title:{display:true,text:"Količina (kom)",font:{family:"IBM Plex Mono",size:9},color:"#8b90a5"}},
          y1:{position:"right",beginAtZero:true,grid:{display:false},ticks:{font:monoFont,color:"#1e3a8a",callback:function(v){return (v/1000).toFixed(0)+"k";}},title:{display:true,text:"Profit (RSD)",font:{family:"IBM Plex Mono",size:9},color:"#1e3a8a"}},
          x:{grid:{display:false},ticks:{font:{family:"IBM Plex Mono",size:8},color:"#8b90a5",maxRotation:45}}
        }
      }
    });
  });
}
var obrtChartInit=false;
function initObrtChart(){
  if(obrtChartInit||typeof Chart==='undefined')return;
  var el=document.getElementById("obrtChart");
  if(!el||!OBRT_DATA||!OBRT_DATA.sistemi||OBRT_DATA.sistemi.length===0)return;
  obrtChartInit=true;
  function colorFor(v){
    if(v<=3) return "#16a34a";
    if(v<=6) return "#0d9488";
    if(v<=12) return "#f59e0b";
    return "#dc2626";
  }
  var colors=OBRT_DATA.meseci.map(colorFor);
  new Chart(el,{
    type:"bar",
    data:{labels:OBRT_DATA.sistemi,datasets:[{data:OBRT_DATA.meseci,backgroundColor:colors,borderRadius:4}]},
    options:{
      responsive:true,maintainAspectRatio:false,
      plugins:{
        legend:{display:false},
        tooltip:{backgroundColor:"#fff",titleColor:"#1a1a2e",bodyColor:"#5a5f7a",borderColor:"#e2e6ef",borderWidth:1,padding:10,titleFont:{family:"IBM Plex Mono",size:11},bodyFont:{family:"IBM Plex Mono",size:10},callbacks:{label:function(c){
          var i=c.dataIndex;
          return [c.parsed.y.toFixed(1)+" meseci za obrt",
                  "Stanje "+OBRT_DATA.last_mes[i]+": "+OBRT_DATA.lager[i].toLocaleString("sr-RS")+" kom",
                  "Ø prodaja: "+OBRT_DATA.prodaja_avg[i].toLocaleString("sr-RS")+"/mes"];
        }}}
      },
      scales:{
        y:{beginAtZero:true,grid:{color:"rgba(0,0,0,0.04)"},ticks:{font:{family:"IBM Plex Mono",size:10},color:"#8b90a5",callback:function(v){return v+" mes";}},title:{display:true,text:"Meseci potrebni za obrt",font:{family:"IBM Plex Mono",size:11},color:"#5a5f7a"}},
        x:{grid:{display:false},ticks:{font:{family:"IBM Plex Mono",size:10,weight:"bold"},color:"#1a1a2e",maxRotation:45,minRotation:45}}
      }
    }
  });
}
var _origShowTab=showTab;
showTab=function(n){_origShowTab(n);if(n==="prodaja")setTimeout(initAnalitikaCharts,50);if(n==="drv")setTimeout(initAkcijaCharts,50);if(n==="zalihe")setTimeout(initObrtChart,50);setTimeout(osveziSkrol,60);};
var SKROL_SYNC=[];
function dupliSkrol(){
  document.querySelectorAll('.tw').forEach(function(tw){
    if(tw.previousElementSibling&&tw.previousElementSibling.classList.contains('tw-top'))return;
    var tbl=tw.querySelector('table'); if(!tbl)return;
    var top=document.createElement('div'); top.className='tw-top';
    var inner=document.createElement('div'); inner.className='tw-top-inner';
    top.appendChild(inner); tw.parentNode.insertBefore(top,tw);
    function sync(){
      var w=tbl.scrollWidth;
      inner.style.width=w+'px';
      top.style.display=(w>tw.clientWidth+2)?'':'none';
    }
    var lock=false;
    top.addEventListener('scroll',function(){if(lock)return;lock=true;tw.scrollLeft=top.scrollLeft;lock=false});
    tw.addEventListener('scroll',function(){if(lock)return;lock=true;top.scrollLeft=tw.scrollLeft;lock=false});
    window.addEventListener('resize',sync);
    SKROL_SYNC.push(sync);
    sync();
  });
}
function osveziSkrol(){SKROL_SYNC.forEach(function(f){try{f()}catch(e){}})}
function initSve(){try{applyFilters()}catch(e){}dupliSkrol();setTimeout(osveziSkrol,150);setTimeout(initAnalitikaCharts,100);}
if(document.readyState!=="loading"){initSve();}
else{document.addEventListener("DOMContentLoaded",initSve);}''')
