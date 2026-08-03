# -*- coding: utf-8 -*-
"""Generisanje direktorskog „Izveštaja prodaje" (dashboard) iz dve tabele.
Isti kod kao samostalna skripta, samo: 2 pitanja su parametri, ulaz su fajlovi
(putanja ili BytesIO), a rezultat je (html_str, xlsx_bytes) umesto pisanja na disk.
"""
import pandas as pd
import json, html as html_mod, io as _io


def _read_excel(src, sheet):
    """Čitanje sheet-a bez obzira da li je src putanja, bytes ili BytesIO (svaki put ispočetka)."""
    if isinstance(src, (bytes, bytearray)):
        return pd.read_excel(_io.BytesIO(src), sheet_name=sheet)
    if hasattr(src, 'seek'):
        try:
            src.seek(0)
        except Exception:
            pass
        return pd.read_excel(src, sheet_name=sheet)
    return pd.read_excel(src, sheet_name=sheet)


def generisi_izvestaj_prodaje(sistemi_src, troskovi_src, potpun=True, iskljuci_poslednji=False):
    file_path = sistemi_src
    troskovi_path = troskovi_src

    df = _read_excel(file_path, 'tabela')
    df.columns = df.columns.astype(str).str.strip()

    df_mag = _read_excel(file_path, 'zalihe magacin ')
    df_mag.columns = df_mag.columns.astype(str).str.strip()
    for c in df_mag.columns:
        if 'kol' in c.lower().strip(): df_mag.rename(columns={c: 'KOL'}, inplace=True); break

    df_troskovi = _read_excel(troskovi_path, 'tabela')
    df_troskovi.columns = df_troskovi.columns.str.strip()
    df_troskovi['SISTEM'] = df_troskovi['SISTEM'].str.strip()
    df_troskovi['Mesec'] = pd.to_numeric(df_troskovi['Mesec'], errors='coerce').fillna(0).astype(int)
    df_troskovi['Godina'] = pd.to_numeric(df_troskovi['Godina'], errors='coerce').fillna(0).astype(int)

    trosak_kolone = [c for c in df_troskovi.columns if c not in ['SISTEM', 'Mesec', 'Godina']]
    trosak_nazivi = {
        'trosak transporta': 'Trošak transporta', 'trosak advokata': 'Trošak advokata',
        'tehnomedia/gigatron': 'Tehnomedia / Gigatron', 'trosak vozila': 'Trošak vozila',
        'reprezentacija': 'Reprezentacija', 'trosak programa': 'Trošak programa',
        'racuni + banka+ osiguranje': 'Računi + banka + osiguranje', 'plate': 'Plate',
        'trosak laboratorije': 'Trošak laboratorije', 'promotivni troskovi': 'Promotivni troškovi',
        'troškovi kanc/ materijala': 'Troškovi kanc. materijala', 'ostali troškovi': 'Ostali troškovi',
    }
    mapa_troskova = {}
    for _, row in df_troskovi.iterrows():
        s = row['SISTEM'];
        g = int(row['Godina']);
        m = int(row['Mesec'])
        key = f"{s}|{g}-{m}";
        mapa_troskova[key] = {}
        for c in trosak_kolone:
            mapa_troskova[key][c] = float(row[c]) if pd.notna(row[c]) else 0

    col_nacin = next((c for c in df.columns if "PLACANJA" in c.upper().replace('Č', 'C').replace('Ć', 'C')), None)
    mapa_placanja = {}
    if col_nacin:
        temp = df[['SISTEM', col_nacin]].dropna()
        mapa_placanja = dict(zip(temp['SISTEM'], temp[col_nacin]))

    df['Mesec'] = pd.to_numeric(df['Mesec'], errors='coerce').fillna(0).astype(int)
    df['Godina'] = pd.to_numeric(df['Godina'], errors='coerce').fillna(0).astype(int)
    df_clean = df[(df['Godina'] >= 2025) & (df['Mesec'] >= 1)].copy()

    col_dod = next((c for c in df_clean.columns if c.strip().upper().startswith('DODATNI MESECNI TROSAK')), None)
    mapa_dodatni = {}
    if col_dod:
        dod_grouped = df_clean.groupby(['SISTEM', 'Godina', 'Mesec'])[col_dod].sum().reset_index()
        for _, row in dod_grouped.iterrows():
            s = row['SISTEM'];
            g = int(row['Godina']);
            m = int(row['Mesec'])
            v = float(row[col_dod]) if pd.notna(row[col_dod]) else 0
            if v > 0:
                mapa_dodatni[f"{s}|{g}-{m}"] = v

    periodi = df_clean[['Godina', 'Mesec']].drop_duplicates().sort_values(['Godina', 'Mesec']).values.tolist()
    mapa_meseci = {1: 'Jan', 2: 'Feb', 3: 'Mar', 4: 'Apr', 5: 'Maj', 6: 'Jun', 7: 'Jul', 8: 'Avg', 9: 'Sep', 10: 'Okt',
                   11: 'Nov', 12: 'Dec'}
    nazivi = [f"{mapa_meseci[int(m)]} {str(int(g))[-2:]}" for g, m in periodi]

    poslednji_g, poslednji_m = periodi[-1]
    poslednji_naziv = f"{mapa_meseci[int(poslednji_m)]} {int(poslednji_g)}"
    if potpun:
        iskljuci = bool(iskljuci_poslednji)
        if iskljuci:
            periodi_profit = periodi[:-1];
            nazivi_profit = nazivi[:-1]
            badge_html = f'<span class="badge br">{poslednji_naziv}: ISKLJUČEN iz profitabilnosti</span>'
        else:
            periodi_profit = periodi[:];
            nazivi_profit = nazivi[:]
            badge_html = f'<span class="badge bg">{poslednji_naziv}: UKLJUČEN</span>'
    else:
        iskljuci = False
        periodi_profit = periodi[:]
        nazivi_profit = nazivi[:]
        badge_html = ''

    periodi_drv = periodi[:]
    nazivi_drv = nazivi[:]

    num_s = len(nazivi)
    num_p = len(nazivi_profit)
    num_d = len(nazivi_drv)

    def get_color(price, grupa=None):
        cm = {1390: '#90EE90', 1300: '#FFD1DC', 1290: '#FFB6C1', 1190: '#FF69B4', 990: '#C71585', 890: '#90EE90',
              800: '#FFD1DC', 790: '#FFB6C1', 730: '#FF69B4', 690: '#C71585'}
        if not price or price == 0: return None
        if abs(price - 590) < 30:
            return '#DC2626' if str(grupa).strip() == 'NERD 2000' else None
        best = min(cm.keys(), key=lambda k: abs(k - price))
        return cm[best] if abs(best - price) < 30 else None

    def esc(t): return html_mod.escape(str(t))

    def fmtnum(v): return f"{round(v):,}"

    def cell_c(bg, val):
        if not bg: return f'<td class="n">{fmtnum(val)}</td>'
        tc = '#fff' if bg in ['#C71585', '#FF69B4', '#DC2626'] else '#1a1a2e'
        return f'<td class="n" style="background:{bg};color:{tc};font-weight:600">{fmtnum(val)}</td>'

    sistemi_lista = sorted(df_clean['SISTEM'].dropna().unique())
    sve_grupe = sorted(df_clean['Grupa artikla'].dropna().astype(str).str.strip().unique())

    prodaja_data_js = {}
    for sistem in sistemi_lista:
        s_data = df_clean[df_clean['SISTEM'] == sistem];
        prodaja_data_js[sistem] = {}
        for grupa in sorted(s_data['Grupa artikla'].dropna().astype(str).str.strip().unique()):
            g_data = s_data[s_data['Grupa artikla'].astype(str).str.strip() == grupa];
            gvals = []
            for g, m in periodi:
                mask = (g_data['Godina'] == int(g)) & (g_data['Mesec'] == int(m))
                gvals.append(round(float(g_data.loc[mask, 'Prodata kolicina ka krajnjem kupcu'].sum())))
            prodaja_data_js[sistem][grupa] = gvals

    profit_data_js = {}
    trosak_names_list = ['Troškovi marketinga']
    trosak_ids_list = ['mkt']
    for ki, kat in enumerate(trosak_kolone):
        trosak_ids_list.append(f"t{ki}");
        trosak_names_list.append(trosak_nazivi.get(kat, kat))
    trosak_ids_list.append('dod')
    trosak_names_list.append('Dodatni mesečni trošak')

    for sistem in sistemi_lista:
        s_data = df_clean[df_clean['SISTEM'] == sistem]
        status = mapa_placanja.get(sistem, None);
        is_f = str(status) in ['1', '1.0']
        pn = [];
        mn = [];
        dn = []
        for g, m in periodi_profit:
            mask = (s_data['Godina'] == int(g)) & (s_data['Mesec'] == int(m))
            pn.append(round(float(s_data.loc[mask, 'PROFIT3' if is_f else 'Profit'].sum())))
            mn.append(round(float(s_data.loc[mask, 'MESECNI TROSAK1'].sum())))
            key = f"{sistem}|{int(g)}-{int(m)}"
            dn.append(round(mapa_dodatni.get(key, 0)))
        profit_data_js[sistem] = {'profit': pn, 'mkt': mn, 'dod': dn}
        for ki, kat in enumerate(trosak_kolone):
            kn = []
            for g, m in periodi_profit:
                key = f"{sistem}|{int(g)}-{int(m)}"
                kn.append(round(mapa_troskova.get(key, {}).get(kat, 0)))
            profit_data_js[sistem][f"t{ki}"] = kn

    drv_data_js = {}
    for sistem in sistemi_lista:
        s_data = df_clean[df_clean['SISTEM'] == sistem];
        pn = [];
        mn = [];
        dn = []
        for g, m in periodi_drv:
            mask = (s_data['Godina'] == int(g)) & (s_data['Mesec'] == int(m))
            pn.append(round(float(s_data.loc[mask, 'Profit'].sum())))
            mn.append(round(float(s_data.loc[mask, 'MESECNI TROSAK1'].sum())))
            key = f"{sistem}|{int(g)}-{int(m)}"
            dn.append(round(mapa_dodatni.get(key, 0)))
        drv_data_js[sistem] = {'profit': pn, 'mkt': mn, 'dod': dn}

    last3 = periodi[-3:]
    df_art = df_clean[df_clean['Artikl'].notna()].copy()
    df_art['Artikl'] = df_art['Artikl'].astype(str).str.strip()
    art_monthly_sales = {}
    for art in df_art['Artikl'].unique():
        a_data = df_art[df_art['Artikl'] == art];
        sales = []
        for g, m in last3:
            mask = (a_data['Godina'] == int(g)) & (a_data['Mesec'] == int(m))
            sales.append(float(a_data.loc[mask, 'Prodata kolicina ka krajnjem kupcu'].sum()))
        art_monthly_sales[art] = round(sum(sales) / len(sales)) if sales else 0

    def get_grupa(art_name):
        if 'HQD' in art_name.upper(): return 'HQD 1000'
        if '2000' in art_name: return 'NERD 2000'
        if 'E-cigareta' in art_name or '1000' in art_name: return 'NERD 1000'
        return 'Ostalo'

    mag_data = [];
    total_mag = 0;
    total_avg = 0
    for _, row in df_mag.iterrows():
        art_name = str(row['Naziv artikla']).strip();
        kol = int(row['KOL'])
        avg_sale = art_monthly_sales.get(art_name, 0)
        daily = avg_sale / 30 if avg_sale > 0 else 0
        days = round(kol / daily) if daily > 0 else 9999
        mag_data.append((art_name, kol, avg_sale, days))
        total_mag += kol;
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
        r = f'<tr class="mag-row" data-grupa="{esc(grupa)}"><td style="font-size:10px;color:var(--t2)">{esc(art_name)}</td><td class="nb" style="color:var(--t2);font-size:10px">{esc(grupa)}</td><td class="nb">{fmtnum(kol)}</td><td class="nb">{fmtnum(avg_sale)}</td><td class="nb" {dcls}>{days_str}</td><td class="nb" {dcls}>{months_str}</td></tr>'
        mag_rows.append(r)

    total_daily = total_avg / 30 if total_avg > 0 else 0
    total_days = round(total_mag / total_daily) if total_daily > 0 else 9999
    total_months_str = f"{total_days / 30:.1f}"
    mag_rows.append(
        f'<tr class="totalrow"><td class="total-label">UKUPNO MAGACIN</td><td></td><td class="nb total-cell">{fmtnum(total_mag)}</td><td class="nb total-cell">{fmtnum(total_avg)}</td><td class="nb total-cell" style="font-weight:800">{total_days}</td><td class="nb total-cell" style="font-weight:800">{total_months_str}</td></tr>')

    mag_grupe_summary = {}
    for art_name, kol, avg_sale, days in mag_data:
        g = get_grupa(art_name)
        if g not in mag_grupe_summary: mag_grupe_summary[g] = {'kol': 0, 'avg': 0}
        mag_grupe_summary[g]['kol'] += kol;
        mag_grupe_summary[g]['avg'] += avg_sale

    def get_last_zalihe(data, periodi_list):
        for g, m in reversed(periodi_list):
            mask = (data['Godina'] == int(g)) & (data['Mesec'] == int(m))
            subset = data.loc[mask];
            zal = subset['Zalihe'].sum()
            if pd.notna(zal) and zal > 0:
                return round(float(zal)), f"{mapa_meseci[int(m)]} {str(int(g))[-2:]}"
        return 0, ""

    zal_rows = [];
    si = 0;
    zal_grand = 0
    for sistem in sistemi_lista:
        si += 1;
        sid = f"zs{si}";
        s_data = df_clean[df_clean['SISTEM'] == sistem]
        s_zal, s_per = get_last_zalihe(s_data, periodi);
        zal_grand += s_zal
        r = f'<tr class="sr" data-sistem="{esc(sistem)}" onclick="tog(\'{sid}\')" style="cursor:pointer"><td><button class="be" id="b-{sid}">+</button><span class="sn">{esc(sistem)}</span></td><td class="nb">{fmtnum(s_zal)}</td><td class="n" style="color:var(--t3);font-size:10px">{s_per}</td></tr>'
        zal_rows.append(r)
        grupe = sorted(s_data['Grupa artikla'].dropna().astype(str).str.strip().unique());
        gi = 0
        for grupa in grupe:
            gi += 1;
            gid = f"{sid}g{gi}";
            g_data = s_data[s_data['Grupa artikla'].astype(str).str.strip() == grupa]
            g_zal, g_per = get_last_zalihe(g_data, periodi)
            r = f'<tr class="gr hidden" data-p="{sid}" data-sistem="{esc(sistem)}" data-grupa="{esc(grupa)}" onclick="tog(\'{gid}\');event.stopPropagation()" style="cursor:pointer"><td style="padding-left:28px"><button class="be beg" id="b-{gid}">+</button><span class="gn">{esc(grupa)}</span></td><td class="nb">{fmtnum(g_zal)}</td><td class="n" style="color:var(--t3);font-size:10px">{g_per}</td></tr>'
            zal_rows.append(r)
            for art in sorted(g_data['Artikl'].dropna().astype(str).str.strip().unique()):
                if not art: continue
                a_data = g_data[g_data['Artikl'].astype(str).str.strip() == art]
                a_zal, a_per = get_last_zalihe(a_data, periodi)
                r = f'<tr class="ar hidden" data-p="{gid}" data-sistem="{esc(sistem)}" data-grupa="{esc(grupa)}"><td class="an">{esc(art)}</td><td class="n" style="color:#7c8494;font-size:10px">{fmtnum(a_zal) if a_zal > 0 else ""}</td><td class="n" style="color:var(--t3);font-size:9px">{a_per}</td></tr>'
                zal_rows.append(r)
        zal_rows.append(f'<tr class="sep" data-sistem="{esc(sistem)}"><td colspan="999"></td></tr>')
    zal_rows.append(
        f'<tr class="totalrow" id="zalihe-total"><td class="total-label">TOTAL</td><td class="nb total-cell" style="font-size:13px">{fmtnum(zal_grand)}</td><td></td></tr>')

    prodaja_rows = [];
    grand_totals = [0] * num_s;
    si = 0
    for sistem in sistemi_lista:
        si += 1;
        sid = f"ps{si}";
        s_data = df_clean[df_clean['SISTEM'] == sistem];
        vals = []
        for g, m in periodi:
            mask = (s_data['Godina'] == int(g)) & (s_data['Mesec'] == int(m))
            vals.append(round(float(s_data.loc[mask, 'Prodata kolicina ka krajnjem kupcu'].sum())))
        total = sum(vals)
        for i in range(num_s): grand_totals[i] += vals[i]
        r = f'<tr class="sr" data-sistem="{esc(sistem)}" onclick="tog(\'{sid}\')" style="cursor:pointer"><td><button class="be" id="b-{sid}">+</button><span class="sn">{esc(sistem)}</span></td>'
        for v in vals: r += f'<td class="nb">{fmtnum(v)}</td>'
        r += f'<td class="nt">{fmtnum(total)}</td></tr>';
        prodaja_rows.append(r)
        grupe = sorted(s_data['Grupa artikla'].dropna().astype(str).str.strip().unique());
        gi = 0
        for grupa in grupe:
            gi += 1;
            gid = f"{sid}g{gi}";
            g_data = s_data[s_data['Grupa artikla'].astype(str).str.strip() == grupa];
            gvals = [];
            gcene = []
            for g, m in periodi:
                mask = (g_data['Godina'] == int(g)) & (g_data['Mesec'] == int(m))
                gvals.append(round(float(g_data.loc[mask, 'Prodata kolicina ka krajnjem kupcu'].sum())))
                ac = g_data.loc[mask, 'FINALNA MP'].mean()
                gcene.append(round(float(ac)) if pd.notna(ac) and ac > 0 else 0)
            gtotal = sum(gvals)
            r = f'<tr class="gr hidden" data-p="{sid}" data-sistem="{esc(sistem)}" data-grupa="{esc(grupa)}" onclick="tog(\'{gid}\');event.stopPropagation()" style="cursor:pointer"><td style="padding-left:28px"><button class="be beg" id="b-{gid}">+</button><span class="gn">{esc(grupa)}</span></td>'
            for i, v in enumerate(gvals): r += cell_c(get_color(gcene[i], grupa), v)
            r += f'<td class="nb">{fmtnum(gtotal)}</td></tr>';
            prodaja_rows.append(r)
            for art in sorted(g_data['Artikl'].dropna().astype(str).str.strip().unique()):
                if not art: continue
                a_data = g_data[g_data['Artikl'].astype(str).str.strip() == art];
                avals = []
                for g, m in periodi:
                    mask = (a_data['Godina'] == int(g)) & (a_data['Mesec'] == int(m))
                    avals.append(round(float(a_data.loc[mask, 'Prodata kolicina ka krajnjem kupcu'].sum())))
                atotal = sum(avals)
                r = f'<tr class="ar hidden" data-p="{gid}" data-sistem="{esc(sistem)}" data-grupa="{esc(grupa)}"><td class="an">{esc(art)}</td>'
                for v in avals:
                    if v > 0:
                        r += f'<td class="n" style="color:#7c8494;font-size:10px">{fmtnum(v)}</td>'
                    else:
                        r += '<td class="n"></td>'
                r += f'<td class="n" style="color:#7c8494;font-size:10px">{fmtnum(atotal)}</td></tr>';
                prodaja_rows.append(r)
        prodaja_rows.append(f'<tr class="sep" data-sistem="{esc(sistem)}"><td colspan="999"></td></tr>')
    gtt = sum(grand_totals)
    tr_r = '<tr class="totalrow" id="prodaja-total"><td class="total-label">TOTAL</td>'
    for v in grand_totals: tr_r += f'<td class="nb total-cell">{fmtnum(v)}</td>'
    tr_r += f'<td class="nt total-cell" style="font-size:13px">{fmtnum(gtt)}</td></tr>';
    prodaja_rows.append(tr_r)

    def pct(curr, ref):
        if ref == 0: return 0.0
        return round((curr - ref) / ref * 100, 1)

    def fmt_diff(v):
        sign = '+' if v >= 0 else ''
        return f"{sign}{fmtnum(v)}"

    last_idx = num_s - 1
    curr_total = grand_totals[last_idx]
    prev_total = grand_totals[last_idx - 1] if num_s >= 2 else 0
    n6 = min(6, last_idx)
    avg6_total = round(sum(grand_totals[last_idx - n6:last_idx]) / n6) if n6 > 0 else 0
    yoy_idx = None
    for idx, (g, m) in enumerate(periodi):
        if int(g) == int(poslednji_g) - 1 and int(m) == int(poslednji_m):
            yoy_idx = idx;
            break
    yoy_total = grand_totals[yoy_idx] if yoy_idx is not None else None

    pct_mom = pct(curr_total, prev_total)
    pct_6m = pct(curr_total, avg6_total)
    pct_yoy = pct(curr_total, yoy_total) if yoy_total else None
    diff_mom = curr_total - prev_total
    diff_6m = curr_total - avg6_total
    diff_yoy = (curr_total - yoy_total) if yoy_total else None

    prev_naziv = nazivi[last_idx - 1] if num_s >= 2 else ""
    yoy_naziv = nazivi[yoy_idx] if yoy_idx is not None else ""

    def badge(v):
        if v is None: return ''
        color = '#16a34a' if v >= 0 else '#dc2626'
        bg = 'rgba(22,163,74,0.08)' if v >= 0 else 'rgba(220,38,38,0.08)'
        arr = '▲' if v >= 0 else '▼'
        sign = 'rast od' if v >= 0 else 'pad od'
        return f'<span style="background:{bg};color:{color};padding:2px 8px;border-radius:6px;font-weight:700;font-family:\'IBM Plex Mono\',monospace;font-size:13px">{arr} {sign} {abs(v)}%</span>'

    headline_parts = []
    headline_parts.append(
        f'Prodaja u <b style="color:#2563eb">{poslednji_naziv.lower()}</b> iznosi <b style="font-family:\'IBM Plex Mono\',monospace;font-size:15px">{fmtnum(curr_total)} kom</b>')
    if num_s >= 2:
        headline_parts.append(
            f'i beleži {badge(pct_mom)} u odnosu na prethodni mesec <span style="color:#8b90a5">({prev_naziv}: {fmtnum(prev_total)} kom)</span>')
    if n6 > 0:
        headline_parts.append(
            f', odnosno {badge(pct_6m)} u odnosu na 6-mesečni prosek <span style="color:#8b90a5">({fmtnum(avg6_total)} kom)</span>')
    if pct_yoy is not None:
        headline_parts.append(
            f'. Poređenjem sa istim mesecom prethodne godine <span style="color:#8b90a5">({yoy_naziv}: {fmtnum(yoy_total)} kom)</span>, rezultat je {badge(pct_yoy)} YoY')
    headline_text = ' '.join(headline_parts) + '.'

    trazene_grupe = ['HQD 1000', 'NERD 2000', 'SYX']
    grupa_boje = {
        'HQD 1000': {'main': '#dc2626', 'bg': 'rgba(220,38,38,0.1)', 'light': 'rgba(220,38,38,0.08)'},
        'NERD 2000': {'main': '#2563eb', 'bg': 'rgba(37,99,235,0.1)', 'light': 'rgba(37,99,235,0.08)'},
        'SYX': {'main': '#16a34a', 'bg': 'rgba(22,163,74,0.1)', 'light': 'rgba(22,163,74,0.08)'},
    }

    grupe_data = {}
    for grupa in trazene_grupe:
        g_data = df_clean[df_clean['Grupa artikla'].astype(str).str.strip() == grupa]
        vals = []
        for g, m in periodi:
            mask = (g_data['Godina'] == int(g)) & (g_data['Mesec'] == int(m))
            vals.append(round(float(g_data.loc[mask, 'Prodata kolicina ka krajnjem kupcu'].sum())))
        grupe_data[grupa] = vals

    total_last_all = curr_total

    grupe_blocks_html = ''
    grupe_chart_data_js = {}

    for grupa in trazene_grupe:
        vals = grupe_data[grupa]
        g_curr = vals[last_idx]
        g_prev = vals[last_idx - 1] if num_s >= 2 else 0
        g_avg6 = round(sum(vals[last_idx - n6:last_idx]) / n6) if n6 > 0 else 0
        g_yoy = vals[yoy_idx] if yoy_idx is not None else None

        g_pct_mom = pct(g_curr, g_prev)
        g_pct_6m = pct(g_curr, g_avg6)
        g_pct_yoy = pct(g_curr, g_yoy) if g_yoy else None

        udeo = round(g_curr / total_last_all * 100, 1) if total_last_all > 0 else 0

        colors = grupa_boje[grupa]
        safe_id = grupa.replace(' ', '_').replace('.', '')

        g_parts = [f'<b style="color:{colors["main"]}">{esc(grupa)}</b> beleži']
        if num_s >= 2:
            g_parts.append(
                f'{badge(g_pct_mom)} u odnosu na prethodni mesec <span style="color:#8b90a5">({prev_naziv}: {fmtnum(g_prev)})</span>,')
        if n6 > 0:
            g_parts.append(f'{badge(g_pct_6m)} u odnosu na 6M prosek <span style="color:#8b90a5">({fmtnum(g_avg6)})</span>')
        if g_pct_yoy is not None:
            g_parts.append(
                f'i {badge(g_pct_yoy)} YoY u odnosu na {yoy_naziv} <span style="color:#8b90a5">({fmtnum(g_yoy)})</span>')
        g_text = ' '.join(g_parts) + '.'

        grupe_chart_data_js[safe_id] = {
            'labels': [f'{poslednji_naziv} (trenutno)', f'{prev_naziv} (-1M)', 'Ø 6 meseci',
                       f'{yoy_naziv} (YoY)' if g_yoy else ''],
            'values': [g_curr, g_prev, g_avg6, g_yoy if g_yoy else 0],
            'has_yoy': g_yoy is not None,
            'main_color': colors['main']
        }

        grupe_blocks_html += f'''
<div class="grupa-block">
  <div class="grupa-header">
    <div class="grupa-title-wrap">
      <span class="grupa-pill" style="background:{colors['bg']};color:{colors['main']}">{esc(grupa)}</span>
      <span class="grupa-udeo">{udeo}% ukupnog miksa</span>
    </div>
    <span class="grupa-kom" style="color:{colors['main']}">{fmtnum(g_curr)} kom</span>
  </div>
  <p class="grupa-text">{g_text}</p>
  <div class="grupa-chart-wrap"><canvas id="chart-{safe_id}"></canvas></div>
</div>
'''

    big_trend_data = {
        'labels': nazivi,
        'values': grand_totals,
        'curr_idx': last_idx,
        'yoy_idx': yoy_idx if yoy_idx is not None else -1
    }

    mini1 = {'labels': [prev_naziv, poslednji_naziv], 'values': [prev_total, curr_total],
             'color': '#dc2626' if pct_mom < 0 else '#16a34a'} if num_s >= 2 else None
    mini2 = {'labels': ['Ø 6M', poslednji_naziv], 'values': [avg6_total, curr_total],
             'color': '#dc2626' if pct_6m < 0 else '#16a34a'} if n6 > 0 else None
    mini3 = {'labels': [yoy_naziv, poslednji_naziv], 'values': [yoy_total, curr_total],
             'color': '#dc2626' if (pct_yoy or 0) < 0 else '#16a34a'} if pct_yoy is not None else None

    analitika_data_json = json.dumps({
        'big': big_trend_data, 'mini1': mini1, 'mini2': mini2, 'mini3': mini3, 'grupe': grupe_chart_data_js
    }, ensure_ascii=False)

    def mini_card(label, pct_val, diff_val, sub_left, sub_right, chart_id):
        if pct_val is None: return ''
        color = '#16a34a' if pct_val >= 0 else '#dc2626'
        sign = '+' if pct_val >= 0 else ''
        return f'''<div class="mini-card" style="border-top:3px solid {color}">
  <div class="mc-lbl">{label}</div>
  <div class="mc-row"><span class="mc-pct" style="color:{color}">{sign}{pct_val}%</span><span class="mc-diff">{fmt_diff(diff_val)} kom</span></div>
  <div class="mc-chart"><canvas id="{chart_id}"></canvas></div>
  <div class="mc-sub"><span>{sub_left}</span><span style="color:{color};font-weight:700">{sub_right}</span></div>
</div>'''

    mini1_html = mini_card('vs PRETHODNI MESEC', pct_mom, diff_mom, f'{prev_naziv}: {fmtnum(prev_total)}',
                           f'{poslednji_naziv}: {fmtnum(curr_total)}', 'mini1') if num_s >= 2 else ''
    mini2_html = mini_card('vs 6M PROSEK', pct_6m, diff_6m, f'Ø 6M: {fmtnum(avg6_total)}',
                           f'{poslednji_naziv}: {fmtnum(curr_total)}', 'mini2') if n6 > 0 else ''
    mini3_html = mini_card(f'YoY ({yoy_naziv} -> {poslednji_naziv})', pct_yoy, diff_yoy,
                           f'{yoy_naziv}: {fmtnum(yoy_total) if yoy_total else "—"}',
                           f'{poslednji_naziv}: {fmtnum(curr_total)}', 'mini3') if pct_yoy is not None else ''

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
            sistem_mom_changes.append({'sistem': sistem, 'pct': pct(curr, prev), 'abs': curr - prev, 'curr': curr, 'prev': prev})

    top3 = sorted([s for s in sistem_mom_changes if s['pct'] > 0], key=lambda x: -x['pct'])[:3]
    bottom3 = sorted([s for s in sistem_mom_changes if s['pct'] < 0], key=lambda x: x['pct'])[:3]

    alarmi = []
    for sistem in sistemi_lista:
        vals = sistem_monthly[sistem]
        curr = vals[last_idx]
        prev = vals[last_idx - 1] if num_s >= 2 else 0
        if num_s >= 4:
            consec_drops = 0
            for i in range(last_idx, 0, -1):
                if vals[i] < vals[i - 1] and vals[i - 1] > 0:
                    consec_drops += 1
                else:
                    break
            if consec_drops >= 3:
                kum_pct = pct(vals[last_idx], vals[last_idx - consec_drops])
                alarmi.append({'level': 'KRITIČNO', 'color': '#dc2626', 'bg': '#fef2f2', 'naziv': sistem, 'tip': 'sistem',
                               'tekst': f'pad {consec_drops} meseca uzastopno ({kum_pct:+}% kumulativno)'})
        if curr == 0 and any(v > 0 for v in vals[:-1]):
            zadnji_aktivan_idx = None
            for i in range(last_idx - 1, -1, -1):
                if vals[i] > 0:
                    zadnji_aktivan_idx = i;
                    break
            if zadnji_aktivan_idx is not None:
                alarmi.append({'level': 'UPOZORENJE', 'color': '#92400e', 'bg': '#fffbeb', 'naziv': sistem, 'tip': 'sistem',
                               'tekst': f'nema prodaje u {poslednji_naziv} (poslednja prodaja: {nazivi[zadnji_aktivan_idx]} — {fmtnum(vals[zadnji_aktivan_idx])} kom)'})
        if curr > 0 and n6 > 0:
            avg6 = round(sum(vals[last_idx - n6:last_idx]) / n6) if n6 > 0 else 0
            if avg6 > 0:
                pct_vs_6m = pct(curr, avg6)
                if pct_vs_6m <= -15:
                    alarmi.append({'level': 'UPOZORENJE', 'color': '#92400e', 'bg': '#fffbeb', 'naziv': sistem, 'tip': 'sistem',
                                   'tekst': f'prodaja ispod 6M proseka za {abs(pct_vs_6m)}% ({fmtnum(curr)} vs Ø {fmtnum(avg6)})'})
        if curr > 0 and curr == max(vals) and curr > sorted(vals)[-2]:
            alarmi.append({'level': 'REKORD', 'color': '#16a34a', 'bg': 'rgba(22,163,74,0.06)', 'naziv': sistem, 'tip': 'sistem',
                           'tekst': f'rekord prodaje u {poslednji_naziv} ({fmtnum(curr)} kom)'})

    glavne_grupe_za_alarme = ['HQD 1000', 'NERD 2000', 'SYX']
    for sistem in sistemi_lista:
        s_data = df_clean[df_clean['SISTEM'] == sistem]
        for grupa in glavne_grupe_za_alarme:
            g_data = s_data[s_data['Grupa artikla'].astype(str).str.strip() == grupa]
            sg_vals = []
            for g, m in periodi:
                mask = (g_data['Godina'] == int(g)) & (g_data['Mesec'] == int(m))
                sg_vals.append(round(float(g_data.loc[mask, 'Prodata kolicina ka krajnjem kupcu'].sum())))
            if sum(sg_vals) == 0: continue
            curr_sg = sg_vals[last_idx]
            if num_s >= 4:
                consec_drops = 0
                for i in range(last_idx, 0, -1):
                    if sg_vals[i] < sg_vals[i - 1] and sg_vals[i - 1] > 0:
                        consec_drops += 1
                    else:
                        break
                if consec_drops >= 3:
                    kum_pct = pct(sg_vals[last_idx], sg_vals[last_idx - consec_drops])
                    alarmi.append({'level': 'KRITIČNO', 'color': '#dc2626', 'bg': '#fef2f2', 'naziv': f'{sistem} / {grupa}', 'tip': 'kombinacija',
                                   'tekst': f'pad {consec_drops} meseca uzastopno ({kum_pct:+}% kumulativno)'})
            if curr_sg == 0 and num_s >= 4:
                recent_avg = sum(sg_vals[max(0, last_idx - 3):last_idx]) / min(3, last_idx)
                if recent_avg >= 50:
                    alarmi.append({'level': 'UPOZORENJE', 'color': '#92400e', 'bg': '#fffbeb', 'naziv': f'{sistem} / {grupa}', 'tip': 'kombinacija',
                                   'tekst': f'nema prodaje u {poslednji_naziv} (prethodno Ø {fmtnum(round(recent_avg))} kom/mesec)'})

    level_order = {'KRITIČNO': 0, 'UPOZORENJE': 1, 'REKORD': 2}
    alarmi.sort(key=lambda x: (level_order.get(x['level'], 99), x['naziv']))
    sistemi_bez_prodaje = set()
    for a in alarmi:
        if a['tip'] == 'sistem' and 'nema prodaje' in a['tekst']:
            sistemi_bez_prodaje.add(a['naziv'])
    sistemi_sa_padom = set()
    for a in alarmi:
        if a['tip'] == 'sistem' and 'pad' in a['tekst'] and 'uzastopno' in a['tekst']:
            sistemi_sa_padom.add(a['naziv'])
    alarmi_filtered = []
    for a in alarmi:
        if a['tip'] == 'kombinacija' and 'nema prodaje' in a['tekst']:
            if a['naziv'].split(' / ')[0] in sistemi_bez_prodaje: continue
        if a['tip'] == 'kombinacija' and 'pad' in a['tekst'] and 'uzastopno' in a['tekst']:
            if a['naziv'].split(' / ')[0] in sistemi_sa_padom: continue
        alarmi_filtered.append(a)
    alarmi = alarmi_filtered

    def top_bottom_row(item, color):
        sign = '+' if item['pct'] >= 0 else ''
        diff_sign = '+' if item['abs'] >= 0 else ''
        return f'''<div class="tb-row">
  <span class="tb-name">{esc(item['sistem'])}</span>
  <span class="tb-vals"><b style="color:{color}">{sign}{item['pct']}%</b> <span class="tb-diff">({diff_sign}{fmtnum(item['abs'])})</span></span>
</div>'''

    top_rows_html = '\n'.join(top_bottom_row(s, '#16a34a') for s in top3) if top3 else '<div class="tb-empty">nema sistema sa rastom</div>'
    bottom_rows_html = '\n'.join(top_bottom_row(s, '#dc2626') for s in bottom3) if bottom3 else '<div class="tb-empty">nema sistema sa padom</div>'

    if alarmi:
        alarmi_html = ''
        for a in alarmi:
            alarmi_html += f'''<div class="alarm-row" style="background:{a['bg']};border-left:3px solid {a['color']}">
  <div class="alarm-text"><b style="color:{a['color']};font-family:'IBM Plex Mono',monospace">{esc(a['naziv'])}</b> <span style="color:var(--t2)">— {esc(a['tekst'])}</span></div>
  <span class="alarm-level" style="color:{a['color']}">{a['level']}</span>
</div>'''
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
    <div class="tb-col tb-col-up"><div class="tb-col-title" style="color:#16a34a">▲ NAJVEĆI RAST</div>{top_rows_html}</div>
    <div class="tb-col tb-col-down"><div class="tb-col-title" style="color:#dc2626">▼ NAJVEĆI PAD</div>{bottom_rows_html}</div>
  </div>
</div>
<div class="alarm-card">
  <div class="alarm-header">
    <h3>⚠️ ALARMI — sistemi i grupe koje zahtevaju pažnju</h3>
    <div class="alarm-counts">
      {'<span class="ac-kritic">●' + str(n_kritic) + ' kritično</span>' if n_kritic > 0 else ''}
      {'<span class="ac-upoz">●' + str(n_upoz) + ' upozorenja</span>' if n_upoz > 0 else ''}
      {'<span class="ac-rekord">●' + str(n_rekord) + ' rekorda</span>' if n_rekord > 0 else ''}
    </div>
  </div>
  <div class="alarm-list">{alarmi_html}</div>
</div>
'''

    analitika_html = f'''
<div class="analitika-wrap">
  <div class="analitika-headline">
    <div class="ah-top"><h2>📊 ANALIZA PRODAJE — {poslednji_naziv.upper()}</h2><span class="ah-badge">automatski ažurirano</span></div>
    <p class="ah-text">{headline_text}</p>
  </div>
  <div class="mini-cards-wrap">{mini1_html}{mini2_html}{mini3_html}</div>
  <div class="big-chart-card">
    <div class="bc-header"><h3>📈 TREND PRODAJE — {num_s} MESECI</h3>
      <div class="bc-legend"><span><span class="lg-line" style="background:#2563eb"></span>Mesečna prodaja</span><span><span class="lg-dash"></span>6M prosek</span><span><span class="lg-dot"></span>Isti mesec prethodne godine</span></div>
    </div>
    <div class="bc-chart"><canvas id="bigTrendChart"></canvas></div>
  </div>
  <div class="grupe-section-header"><h2>📦 ANALIZA PO GRUPAMA — {poslednji_naziv.upper()}</h2><div class="gsh-sub">Detaljna razrada za HQD 1000, NERD 2000 i SYX</div></div>
  {grupe_blocks_html}
  {top_bottom_html}
</div>
'''

    # === PROFIT ROWS ===
    profit_rows = [];
    si = 0;
    profit_grand_totals = [0] * num_p
    for sistem in sistemi_lista:
        si += 1;
        sid2 = f"pr{si}";
        pd_s = profit_data_js[sistem]
        status = mapa_placanja.get(sistem, None);
        is_f = str(status) in ['1', '1.0'];
        is_o = str(status) in ['0', '0.0']
        profit_niz = pd_s['profit'];
        mkt_niz = pd_s['mkt'];
        dod_niz = pd_s['dod']
        uk_niz = list(mkt_niz)
        for tid in trosak_ids_list[1:]:
            if tid == 'dod':
                for j in range(num_p): uk_niz[j] += dod_niz[j]
            else:
                for j in range(num_p): uk_niz[j] += pd_s.get(tid, [0] * num_p)[j]
        neto_niz = [profit_niz[j] - uk_niz[j] for j in range(num_p)]
        for j in range(num_p): profit_grand_totals[j] += neto_niz[j]
        nacin = "po fakturi" if is_f else ("po odjavi" if is_o else "—");
        ncls = "nf" if is_f else "no"
        r = f'<tr class="sr profit-sr" data-sistem="{esc(sistem)}" data-sid="{sid2}" onclick="tog(\'{sid2}\')" style="cursor:pointer"><td><button class="be" id="b-{sid2}">+</button><span class="sn">{esc(sistem)}</span></td><td class="{ncls}">{nacin}</td><td class="nb">Neto profit</td>'
        for v in neto_niz:
            cls = "np" if v >= 0 else "nn";
            r += f'<td class="nb {cls}">{fmtnum(v)}</td>'
        neto_total = sum(neto_niz);
        ncl = "np" if neto_total >= 0 else "nn"
        r += f'<td class="nb {ncl}" style="font-size:13px">{fmtnum(neto_total)}</td></tr>';
        profit_rows.append(r)
        stavka = "Profit (promet)" if is_f else ("Profit (odjava)" if is_o else "—")
        r = f'<tr class="cr hidden" data-p="{sid2}" data-sistem="{esc(sistem)}"><td></td><td></td><td class="cl" style="color:var(--t2);font-style:normal;font-weight:600">{stavka}</td>'
        for v in profit_niz: r += f'<td class="n" style="font-weight:600;color:var(--t2)">{fmtnum(v)}</td>'
        r += f'<td class="n" style="font-weight:700;color:var(--t2)">{fmtnum(sum(profit_niz))}</td></tr>';
        profit_rows.append(r)
        r = f'<tr class="cr hidden" data-p="{sid2}" data-sistem="{esc(sistem)}" data-trosak="mkt"><td></td><td></td><td class="cl">Troškovi marketinga</td>'
        for v in mkt_niz: r += f'<td class="n cc">{fmtnum(v)}</td>'
        r += f'<td class="n cc" style="font-weight:700">{fmtnum(sum(mkt_niz))}</td></tr>';
        profit_rows.append(r)
        for ki, kat in enumerate(trosak_kolone):
            tid = f"t{ki}";
            vals = pd_s.get(tid, [0] * num_p);
            naz = trosak_nazivi.get(kat, kat)
            r = f'<tr class="cr hidden" data-p="{sid2}" data-sistem="{esc(sistem)}" data-trosak="{tid}"><td></td><td></td><td class="cl">{esc(naz)}</td>'
            for v in vals: r += f'<td class="n cc">{fmtnum(v)}</td>'
            r += f'<td class="n cc" style="font-weight:700">{fmtnum(sum(vals))}</td></tr>';
            profit_rows.append(r)
        r = f'<tr class="cr hidden" data-p="{sid2}" data-sistem="{esc(sistem)}" data-trosak="dod"><td></td><td></td><td class="cl">Dodatni mesečni trošak</td>'
        for v in dod_niz: r += f'<td class="n cc">{fmtnum(v)}</td>'
        r += f'<td class="n cc" style="font-weight:700">{fmtnum(sum(dod_niz))}</td></tr>';
        profit_rows.append(r)
        r = f'<tr class="ctr hidden ukupni-row" data-p="{sid2}" data-sistem="{esc(sistem)}" data-sid="{sid2}"><td></td><td></td><td class="ctl">UKUPNI TROŠKOVI</td>'
        for v in uk_niz: r += f'<td class="n ctc">{fmtnum(v)}</td>'
        r += f'<td class="n ctc" style="font-size:12px">{fmtnum(sum(uk_niz))}</td></tr>';
        profit_rows.append(r)
        r = f'<tr class="nr hidden neto-row" data-p="{sid2}" data-sistem="{esc(sistem)}" data-sid="{sid2}"><td></td><td></td><td class="nl">NETO PROFIT</td>'
        for v in neto_niz:
            cls = "np" if v >= 0 else "nn";
            r += f'<td class="n {cls}">{fmtnum(v)}</td>'
        r += f'<td class="n {ncl}" style="font-size:13px">{fmtnum(neto_total)}</td></tr>';
        profit_rows.append(r)
        profit_rows.append(f'<tr class="sep" data-sistem="{esc(sistem)}"><td colspan="999"></td></tr>')
    pgt = sum(profit_grand_totals);
    pgcl = "np" if pgt >= 0 else "nn"
    tr_p = '<tr class="totalrow" id="profit-total"><td class="total-label">TOTAL</td><td></td><td></td>'
    for v in profit_grand_totals:
        cls = "np" if v >= 0 else "nn";
        tr_p += f'<td class="nb total-cell {cls}">{fmtnum(v)}</td>'
    tr_p += f'<td class="nb total-cell {pgcl}" style="font-size:13px">{fmtnum(pgt)}</td></tr>';
    profit_rows.append(tr_p)

    # === DR VUKAŠIN ROWS ===
    drv_rows = [];
    si = 0;
    drv_grand_totals = [0] * num_d
    for sistem in sistemi_lista:
        si += 1;
        sid3 = f"dv{si}";
        dv = drv_data_js[sistem]
        profit_niz = dv['profit'];
        mkt_niz = dv['mkt'];
        dod_niz = dv['dod']
        uk_niz = [mkt_niz[j] + dod_niz[j] for j in range(num_d)]
        neto_niz = [profit_niz[j] - uk_niz[j] for j in range(num_d)]
        for j in range(num_d): drv_grand_totals[j] += neto_niz[j]
        r = f'<tr class="sr" data-sistem="{esc(sistem)}" data-sid="{sid3}" onclick="tog(\'{sid3}\')" style="cursor:pointer"><td><button class="be" id="b-{sid3}">+</button><span class="sn">{esc(sistem)}</span></td><td class="no">po odjavi</td><td class="nb">Neto profit</td>'
        for v in neto_niz:
            cls = "np" if v >= 0 else "nn";
            r += f'<td class="nb {cls}">{fmtnum(v)}</td>'
        neto_total = sum(neto_niz);
        ncl = "np" if neto_total >= 0 else "nn"
        r += f'<td class="nb {ncl}" style="font-size:13px">{fmtnum(neto_total)}</td></tr>';
        drv_rows.append(r)
        r = f'<tr class="cr hidden" data-p="{sid3}" data-sistem="{esc(sistem)}"><td></td><td></td><td class="cl" style="color:var(--t2);font-style:normal;font-weight:600">Profit (odjava)</td>'
        for v in profit_niz: r += f'<td class="n" style="font-weight:600;color:var(--t2)">{fmtnum(v)}</td>'
        r += f'<td class="n" style="font-weight:700;color:var(--t2)">{fmtnum(sum(profit_niz))}</td></tr>';
        drv_rows.append(r)
        r = f'<tr class="cr hidden" data-p="{sid3}" data-sistem="{esc(sistem)}" data-trosak="mkt"><td></td><td></td><td class="cl">Troškovi marketinga</td>'
        for v in mkt_niz: r += f'<td class="n cc">{fmtnum(v)}</td>'
        r += f'<td class="n cc" style="font-weight:700">{fmtnum(sum(mkt_niz))}</td></tr>';
        drv_rows.append(r)
        r = f'<tr class="cr hidden" data-p="{sid3}" data-sistem="{esc(sistem)}" data-trosak="dod"><td></td><td></td><td class="cl">Dodatni mesečni trošak</td>'
        for v in dod_niz: r += f'<td class="n cc">{fmtnum(v)}</td>'
        r += f'<td class="n cc" style="font-weight:700">{fmtnum(sum(dod_niz))}</td></tr>';
        drv_rows.append(r)
        r = f'<tr class="ctr hidden" data-p="{sid3}" data-sistem="{esc(sistem)}" data-sid="{sid3}"><td></td><td></td><td class="ctl">UKUPNI TROŠKOVI</td>'
        for v in uk_niz: r += f'<td class="n ctc">{fmtnum(v)}</td>'
        r += f'<td class="n ctc" style="font-size:12px">{fmtnum(sum(uk_niz))}</td></tr>';
        drv_rows.append(r)
        r = f'<tr class="nr hidden neto-row-dv" data-p="{sid3}" data-sistem="{esc(sistem)}" data-sid="{sid3}"><td></td><td></td><td class="nl">NETO PROFIT</td>'
        for v in neto_niz:
            cls = "np" if v >= 0 else "nn";
            r += f'<td class="n {cls}">{fmtnum(v)}</td>'
        r += f'<td class="n {ncl}" style="font-size:13px">{fmtnum(neto_total)}</td></tr>';
        drv_rows.append(r)
        drv_rows.append(f'<tr class="sep" data-sistem="{esc(sistem)}"><td colspan="999"></td></tr>')
    dgt = sum(drv_grand_totals);
    dgcl = "np" if dgt >= 0 else "nn"
    tr_d = '<tr class="totalrow" id="drv-total"><td class="total-label">TOTAL</td><td></td><td></td>'
    for v in drv_grand_totals:
        cls = "np" if v >= 0 else "nn";
        tr_d += f'<td class="nb total-cell {cls}">{fmtnum(v)}</td>'
    tr_d += f'<td class="nb total-cell {dgcl}" style="font-size:13px">{fmtnum(dgt)}</td></tr>';
    drv_rows.append(tr_d)

    # === ANALIZA USPEŠNOSTI AKCIJE (skraćeno na neophodno) ===
    df_art_full = df_clean[df_clean['Artikl'].notna()].copy()
    df_art_full['ima_akciju'] = df_art_full['AKCIJSKE CENE'].fillna(0) > 0
    POSMATRANE_GRUPE = ['HQD 1000', 'NERD 2000']
    akcija_blokovi_data = []

    def bruto_profit_artikli_mesec(s_art_df, artikli, godina, mesec):
        mask = (s_art_df['Godina'] == int(godina)) & (s_art_df['Mesec'] == int(mesec)) & (s_art_df['Artikl'].isin(artikli))
        if not mask.any(): return 0
        return float(s_art_df.loc[mask, 'Profit'].sum())

    for sistem in sistemi_lista:
        s_art_all = df_art_full[df_art_full['SISTEM'] == sistem].copy()
        if len(s_art_all) == 0: continue
        s_art_all['Grupa_clean'] = s_art_all['Grupa artikla'].astype(str).str.strip()
        s_art = s_art_all[s_art_all['Grupa_clean'].isin(POSMATRANE_GRUPE)].copy()
        if len(s_art) == 0: continue
        g_per_mes = s_art.groupby(['Godina', 'Mesec']).agg(na_akciji=('ima_akciju', 'sum')).reset_index()
        g_per_mes['ima_neku_akciju'] = g_per_mes['na_akciji'] > 0
        g_per_mes = g_per_mes.sort_values(['Godina', 'Mesec']).reset_index(drop=True)
        akcijski_meseci_df = g_per_mes[g_per_mes['ima_neku_akciju']]
        if len(akcijski_meseci_df) == 0: continue
        last_akc_row = akcijski_meseci_df.iloc[-1]
        last_akc_g = int(last_akc_row['Godina']);
        last_akc_m = int(last_akc_row['Mesec'])
        redovni_meseci_df = g_per_mes[~g_per_mes['ima_neku_akciju']]
        if len(redovni_meseci_df) == 0: continue
        last_neakc_row = redovni_meseci_df.iloc[-1]
        last_neakc_g = int(last_neakc_row['Godina']);
        last_neakc_m = int(last_neakc_row['Mesec'])
        akc_data = s_art[(s_art['Godina'] == last_akc_g) & (s_art['Mesec'] == last_akc_m) & (s_art['ima_akciju'])]
        if len(akc_data) == 0: continue
        akcijski_artikli = akc_data['Artikl'].unique().tolist()
        neakc_data = s_art[(s_art['Godina'] == last_neakc_g) & (s_art['Mesec'] == last_neakc_m) & (s_art['Artikl'].isin(akcijski_artikli))]
        grupe_u_akciji = sorted(akc_data['Grupa artikla'].dropna().astype(str).str.strip().unique())
        grupe_breakdown = [];
        total_akc_qty_komparabilno = 0;
        total_neakc_qty_komparabilno = 0;
        komparabilne_grupe = []
        for grupa in grupe_u_akciji:
            akc_g = akc_data[akc_data['Grupa artikla'].astype(str).str.strip() == grupa]
            neakc_g = neakc_data[neakc_data['Grupa artikla'].astype(str).str.strip() == grupa]
            akc_g_cena = float(akc_g['AKCIJSKE CENE'].mean()) if len(akc_g) > 0 else 0
            neakc_g_cena_ref = float(akc_g['Redovna MP CENA'].mean()) if len(akc_g) > 0 else 0
            neakc_g_cena_stvarna = float(neakc_g['Redovna MP CENA'].mean()) if len(neakc_g) > 0 else 0
            neakc_g_cena = neakc_g_cena_stvarna if neakc_g_cena_stvarna > 0 else neakc_g_cena_ref
            akc_g_qty = float(akc_g['Prodata kolicina ka krajnjem kupcu'].sum())
            neakc_g_qty = float(neakc_g['Prodata kolicina ka krajnjem kupcu'].sum()) if len(neakc_g) > 0 else 0
            has_neakc_data = len(neakc_g) > 0
            if has_neakc_data:
                total_akc_qty_komparabilno += akc_g_qty;
                total_neakc_qty_komparabilno += neakc_g_qty;
                komparabilne_grupe.append(grupa)
            grupe_breakdown.append({'grupa': grupa, 'akc_cena': akc_g_cena, 'neakc_cena': neakc_g_cena,
                                    'akc_qty': akc_g_qty, 'neakc_qty': neakc_g_qty, 'has_neakc_data': has_neakc_data,
                                    'broj_artikala': len(akc_g)})
        if len(komparabilne_grupe) == 0: continue
        akcijski_artikli_komparabilni = akc_data[akc_data['Grupa artikla'].astype(str).str.strip().isin(komparabilne_grupe)]['Artikl'].unique().tolist()
        broj_artikala_komparabilni = len(akcijski_artikli_komparabilni)
        akc_bruto = bruto_profit_artikli_mesec(s_art, akcijski_artikli_komparabilni, last_akc_g, last_akc_m)
        neakc_bruto = bruto_profit_artikli_mesec(s_art, akcijski_artikli_komparabilni, last_neakc_g, last_neakc_m)
        pct_qty_total = pct(total_akc_qty_komparabilno, total_neakc_qty_komparabilno) if total_neakc_qty_komparabilno > 0 else None
        pct_bruto = pct(akc_bruto, neakc_bruto) if neakc_bruto != 0 else None
        if pct_bruto is None:
            status_label = '— nema reference'; status_color = '#8b90a5'; status_bg = 'rgba(139,144,165,0.08)'
            zaklj_border = '#8b90a5'; zaklj_bg_grad = 'rgba(139,144,165,0.05),rgba(0,0,0,0.02)'; zaklj_color = '#5a5f7a'
        elif pct_bruto >= 10:
            status_label = '✓ ISPLATILA SE'; status_color = '#16a34a'; status_bg = 'rgba(22,163,74,0.08)'
            zaklj_border = '#16a34a'; zaklj_bg_grad = 'rgba(22,163,74,0.05),rgba(37,99,235,0.03)'; zaklj_color = '#16a34a'
        elif pct_bruto >= -10:
            status_label = '~ MARGINALNO'; status_color = '#92400e'; status_bg = 'rgba(245,158,11,0.08)'
            zaklj_border = '#f59e0b'; zaklj_bg_grad = 'rgba(245,158,11,0.05),rgba(220,38,38,0.03)'; zaklj_color = '#92400e'
        else:
            status_label = '✗ NIJE SE ISPLATILA'; status_color = '#dc2626'; status_bg = 'rgba(220,38,38,0.08)'
            zaklj_border = '#dc2626'; zaklj_bg_grad = 'rgba(220,38,38,0.05),rgba(245,158,11,0.03)'; zaklj_color = '#dc2626'
        last_akc_naziv = f"{mapa_meseci[last_akc_m]} {str(last_akc_g)[-2:]}"
        last_neakc_naziv = f"{mapa_meseci[last_neakc_m]} {str(last_neakc_g)[-2:]}"
        monthly_qty = [];
        monthly_profit = [];
        monthly_is_akcija = []
        for gper, mper in periodi:
            m_data = s_art[(s_art['Godina'] == int(gper)) & (s_art['Mesec'] == int(mper)) & (s_art['Artikl'].isin(akcijski_artikli_komparabilni))]
            monthly_qty.append(round(float(m_data['Prodata kolicina ka krajnjem kupcu'].sum())))
            monthly_profit.append(round(float(m_data['Profit'].sum())))
            monthly_is_akcija.append(bool((m_data['ima_akciju'] == True).any()))
        akcija_blokovi_data.append({
            'sistem': sistem, 'broj_artikala': broj_artikala_komparabilni, 'last_akc_naziv': last_akc_naziv,
            'last_neakc_naziv': last_neakc_naziv, 'grupe_breakdown': grupe_breakdown,
            'total_akc_qty': total_akc_qty_komparabilno, 'total_neakc_qty': total_neakc_qty_komparabilno,
            'pct_qty_total': pct_qty_total, 'akc_bruto': akc_bruto, 'neakc_bruto': neakc_bruto, 'pct_bruto': pct_bruto,
            'status_label': status_label, 'status_color': status_color, 'status_bg': status_bg,
            'zaklj_border': zaklj_border, 'zaklj_bg_grad': zaklj_bg_grad, 'zaklj_color': zaklj_color,
            'monthly_qty': monthly_qty, 'monthly_profit': monthly_profit, 'monthly_is_akcija': monthly_is_akcija})

    akcija_blokovi_data.sort(key=lambda x: (9999 if x['pct_bruto'] is None else -x['pct_bruto']))

    def fmt_pct_signed(v):
        if v is None: return "—"
        return f"{'+' if v >= 0 else ''}{v}%"

    def color_pct(v):
        if v is None: return '#8b90a5'
        return '#16a34a' if v > 0 else ('#dc2626' if v < 0 else '#8b90a5')

    akcija_html_blokovi = ''
    for d in akcija_blokovi_data:
        safe_id = d['sistem'].replace(' ', '_').replace('/', '_').replace('.', '')
        grupe_rows_parts = []
        for gb in d['grupe_breakdown']:
            pct_cena_g = pct(gb['akc_cena'], gb['neakc_cena']) if gb['neakc_cena'] > 0 else None
            pct_qty_g = pct(gb['akc_qty'], gb['neakc_qty']) if gb['neakc_qty'] > 0 else None
            if gb['has_neakc_data']:
                qty_neakc_cell = fmtnum(gb['neakc_qty'])
                pct_qty_cell = f'<span style="color:{color_pct(pct_qty_g)};font-weight:700">{fmt_pct_signed(pct_qty_g)}</span>'
            else:
                qty_neakc_cell = '<span style="color:var(--t3);font-style:italic">nema</span>'
                pct_qty_cell = '<span style="color:var(--t3)">—</span>'
            grupe_rows_parts.append(
                '<tr>'
                f'<td class="g-naziv">{esc(gb["grupa"])} <span class="g-broj">({gb["broj_artikala"]} art.)</span></td>'
                f'<td class="g-num">{fmtnum(gb["akc_cena"])} RSD</td><td class="g-num">{fmtnum(gb["neakc_cena"])} RSD</td>'
                f'<td class="g-num" style="color:{color_pct(pct_cena_g)}">{fmt_pct_signed(pct_cena_g)}</td>'
                f'<td class="g-num" style="color:var(--grn);font-weight:700">{fmtnum(gb["akc_qty"])}</td>'
                f'<td class="g-num">{qty_neakc_cell}</td><td class="g-num">{pct_qty_cell}</td></tr>')
        grupe_rows_html = ''.join(grupe_rows_parts)
        if d['pct_bruto'] is None:
            zakljucak_tekst = (f'Za <b>{d["broj_artikala"]} artikala</b> na akciji ({esc(d["last_akc_naziv"])} vs {esc(d["last_neakc_naziv"])}): '
                               f'prodato <b style="color:var(--grn)">{fmtnum(d["total_akc_qty"])} kom</b>. Nema reference za profit.')
        else:
            if d['pct_bruto'] >= 10:
                komentar = "Lager očišćen efikasno, profit veći uprkos nižoj marži."
            elif d['pct_bruto'] >= -10:
                komentar = "Efekat neutralan — povećana prodaja kompenzuje sniženu maržu."
            else:
                komentar = "Razmotri da li je popust preagresivan — gubitak marže veći od dobitka u količini."
            qty_part = ''
            if d['pct_qty_total'] is not None:
                qty_part = (f'Prodaja je <b style="color:{color_pct(d["pct_qty_total"])}">{fmt_pct_signed(d["pct_qty_total"])}</b> '
                            f'<span style="color:var(--t3)">({fmtnum(d["total_akc_qty"])} vs {fmtnum(d["total_neakc_qty"])} kom)</span>. ')
            zakljucak_tekst = (
                f'Za <b>{d["broj_artikala"]} artikala</b> na akciji, poređenjem <b>{esc(d["last_akc_naziv"])}</b> i <b>{esc(d["last_neakc_naziv"])}</b>: {qty_part}<br>'
                f'<span style="font-family:\'IBM Plex Mono\',monospace;font-size:12px">Bruto profit:</span> '
                f'<b style="color:{color_pct(d["pct_bruto"])};font-family:\'IBM Plex Mono\',monospace;font-size:13px">{fmt_pct_signed(d["pct_bruto"])}</b> '
                f'<span style="color:var(--t3)">({fmtnum(d["akc_bruto"])} vs {fmtnum(d["neakc_bruto"])} RSD)</span>. {komentar}')
        akcija_html_blokovi += (
            '<div class="ak-block"><div class="ak-block-header"><div class="ak-sistem-info">'
            f'<span class="ak-pill">{esc(d["sistem"])}</span>'
            f'<span class="ak-meta">{d["broj_artikala"]} artikala · {esc(d["last_akc_naziv"])} (akcija) vs {esc(d["last_neakc_naziv"])} (redovno)</span></div>'
            f'<span class="ak-status" style="background:{d["status_bg"]};color:{d["status_color"]}">{d["status_label"]}</span></div>'
            f'<div class="ak-zakljucak" style="background:linear-gradient(90deg,{d["zaklj_bg_grad"]});border-left:3px solid {d["zaklj_border"]}">'
            f'<div class="ak-zakljucak-label" style="color:{d["zaklj_color"]}">📌 ZAKLJUČAK</div>'
            f'<p style="font-size:13px;line-height:1.7;color:var(--t1);margin:0">{zakljucak_tekst}</p></div>'
            '<div class="ak-grupe-table-wrap"><div class="ak-table-title">💰 Cene i količine — po grupama</div>'
            '<table class="ak-grupe-table"><thead><tr><th rowspan="2" style="text-align:left">GRUPA</th>'
            '<th colspan="3" class="th-cena">CENA (RSD)</th><th colspan="3" class="th-qty">KOLIČINA (kom)</th></tr>'
            '<tr><th class="th-akc">akcija</th><th class="th-neakc">redovno</th><th class="th-raz">razlika</th>'
            '<th class="th-akc">akcija</th><th class="th-neakc">redovno</th><th class="th-raz">razlika</th></tr></thead>'
            f'<tbody>{grupe_rows_html}</tbody></table></div>'
            '<div class="ak-neto-row">'
            f'<div class="ak-neto-box ak-neto-akc"><div class="ak-neto-lbl">BRUTO PROFIT — {esc(d["last_akc_naziv"])} (akcija)</div><div class="ak-neto-val">{fmtnum(d["akc_bruto"])} RSD</div></div>'
            f'<div class="ak-neto-box ak-neto-neakc"><div class="ak-neto-lbl">BRUTO PROFIT — {esc(d["last_neakc_naziv"])} (redovno)</div><div class="ak-neto-val">{fmtnum(d["neakc_bruto"])} RSD</div></div>'
            f'<div class="ak-neto-box ak-neto-razlika"><div class="ak-neto-lbl">RAZLIKA</div><div class="ak-neto-val" style="color:{color_pct(d["pct_bruto"])}">{fmt_pct_signed(d["pct_bruto"])}</div></div></div>'
            '<div class="ak-chart-wrap"><div class="ak-chart-legend"><span class="ak-chart-title">📈 Prodaja + bruto profit po mesecu</span>'
            '<span class="ak-legend-items"><span><span class="ak-leg-sq ak-leg-akcija"></span>akcija</span>'
            '<span><span class="ak-leg-sq ak-leg-redovno"></span>redovno</span><span><span class="ak-leg-profit"></span>bruto profit</span></span></div>'
            f'<div class="ak-chart"><canvas id="ak-chart-{safe_id}"></canvas></div></div></div>')

    akcija_chart_data = {}
    for d in akcija_blokovi_data:
        safe_id = d['sistem'].replace(' ', '_').replace('/', '_').replace('.', '')
        akcija_chart_data[safe_id] = {'labels': nazivi, 'values': d['monthly_qty'],
                                      'profit': d['monthly_profit'], 'is_akcija': d['monthly_is_akcija']}
    akcija_data_json = json.dumps(akcija_chart_data, ensure_ascii=False)

    if akcija_blokovi_data:
        akcija_analiza_html = ('<div class="akcija-analiza-wrap"><div class="aa-header"><h2>🎯 ANALIZA USPEŠNOSTI AKCIJE — PO SISTEMIMA</h2>'
                               '<div class="aa-sub">Poređenje poslednjeg akcijskog i najbližeg neakcijskog meseca — samo akcijski artikli. Sortirano po profitabilnosti.</div></div>'
                               + akcija_html_blokovi + '</div>')
    else:
        akcija_analiza_html = '<div style="padding:20px;text-align:center;color:var(--t3);font-family:\'IBM Plex Mono\',monospace;font-size:12px">Nema dovoljno podataka za analizu akcije.</div>'

    # === OBRT LAGERA ===
    obrt_lagera_data = []
    for sistem in sistemi_lista:
        s_data = df_clean[df_clean['SISTEM'] == sistem]
        if len(s_data) == 0: continue
        s_zal_per_mes = s_data.groupby(['Godina', 'Mesec']).agg(zal=('Zalihe', 'sum')).reset_index()
        s_zal_per_mes = s_zal_per_mes[s_zal_per_mes['zal'] > 0].sort_values(['Godina', 'Mesec'])
        if len(s_zal_per_mes) == 0: continue
        last_row = s_zal_per_mes.iloc[-1]
        sys_last_g = int(last_row['Godina']);
        sys_last_m = int(last_row['Mesec']);
        trenutni_lager = float(last_row['zal'])
        idx_in_periodi = None
        for i, (g, m) in enumerate(periodi):
            if int(g) == sys_last_g and int(m) == sys_last_m:
                idx_in_periodi = i;
                break
        if idx_in_periodi is None: continue
        start_idx = max(0, idx_in_periodi - 2)
        last_3_sys = periodi[start_idx:idx_in_periodi + 1]
        prodaja_sum = 0
        for g3, m3 in last_3_sys:
            mask3 = (s_data['Godina'] == int(g3)) & (s_data['Mesec'] == int(m3))
            prodaja_sum += float(s_data.loc[mask3, 'Prodata kolicina ka krajnjem kupcu'].sum())
        prosecna_prodaja = prodaja_sum / len(last_3_sys) if len(last_3_sys) > 0 else 0
        if prosecna_prodaja <= 0: continue
        obrt_lagera_data.append({'sistem': sistem, 'lager': trenutni_lager, 'prodaja_avg': prosecna_prodaja,
                                 'meseci': trenutni_lager / prosecna_prodaja, 'last_mes_naziv': f"{mapa_meseci[sys_last_m]} {str(sys_last_g)[-2:]}"})
    obrt_lagera_data.sort(key=lambda x: x['meseci'])
    obrt_chart_json = json.dumps({
        'sistemi': [d['sistem'] for d in obrt_lagera_data], 'meseci': [round(d['meseci'], 1) for d in obrt_lagera_data],
        'lager': [round(d['lager']) for d in obrt_lagera_data], 'prodaja_avg': [round(d['prodaja_avg']) for d in obrt_lagera_data],
        'last_mes': [d['last_mes_naziv'] for d in obrt_lagera_data]}, ensure_ascii=False)

    if obrt_lagera_data:
        obrt_lagera_html = '''
<div class="obrt-wrap"><div class="obrt-header"><h3>⚙️ OBRT LAGERA PO SISTEMIMA</h3>
<div class="obrt-sub">Trenutni lager / Ø prodaja poslednja 3 meseca = broj meseci za obrt</div></div>
<div class="obrt-card"><div class="obrt-legend-row"><span class="obrt-title">📊 BROJ MESECI ZA OBRT LAGERA</span>
<span class="obrt-legend"><span><span class="obrt-leg-sq" style="background:#16a34a"></span>≤3 mes</span>
<span><span class="obrt-leg-sq" style="background:#0d9488"></span>3-6</span><span><span class="obrt-leg-sq" style="background:#f59e0b"></span>6-12</span>
<span><span class="obrt-leg-sq" style="background:#dc2626"></span>&gt;12</span></span></div>
<div class="obrt-chart"><canvas id="obrtChart"></canvas></div>
<div class="obrt-note">Sortirano od najbržeg ka najsporijem obrtu.</div></div></div>'''
    else:
        obrt_lagera_html = ''

    # === HEADERS ===
    ph = '<tr><th style="text-align:left;min-width:280px">SISTEM / GRUPA / ARTIKAL</th>'
    for n in nazivi: ph += f'<th>{n}</th>'
    ph += '<th>TOTAL</th></tr>'
    prh = '<tr><th style="text-align:left;min-width:200px">SISTEM</th><th>NAČIN</th><th>STAVKA</th>'
    for n in nazivi_profit: prh += f'<th>{n}</th>'
    prh += '<th>TOTAL</th></tr>'
    drvh = '<tr><th style="text-align:left;min-width:200px">SISTEM</th><th>NAČIN</th><th>STAVKA</th>'
    for n in nazivi_drv: drvh += f'<th>{n}</th>'
    drvh += '<th>TOTAL</th></tr>'
    mh = '<tr><th style="text-align:left;min-width:300px">ARTIKAL</th><th>GRUPA</th><th>MAGACIN (kom)</th><th>Ø PRODAJA/mes</th><th>DANA ZALIHA</th><th>MESECI</th></tr>'
    zsh = '<tr><th style="text-align:left;min-width:280px">SISTEM / GRUPA / ARTIKAL</th><th>ZALIHE (kom)</th><th>PERIOD</th></tr>'

    info = f"{len(sistemi_lista)} sistema · {len(nazivi)} meseci"
    sistem_options = ''.join([f'<option value="{esc(s)}">{esc(s)}</option>' for s in sistemi_lista])
    grupa_options = ''.join([f'<option value="{esc(g)}">{esc(g)}</option>' for g in sve_grupe])
    trosak_checks = ''
    for tid, tname in zip(trosak_ids_list, trosak_names_list):
        trosak_checks += f'<label class="tcb"><input type="checkbox" checked value="{tid}" onchange="applyProfitFilters()"><span>{tname}</span></label>\n    '

    prodaja_json = json.dumps(prodaja_data_js, ensure_ascii=False)
    profit_json = json.dumps(profit_data_js, ensure_ascii=False)
    drv_json = json.dumps(drv_data_js, ensure_ascii=False)

    last3_names = [nazivi[i] for i in range(-3, 0)]
    mag_info = f"Prosek prodaje: {', '.join(last3_names)}"

    cards_html = f'''<div class="mag-cards">
  <div class="mag-card"><div class="mc-label">Ukupno magacin</div><div class="mc-val" style="color:var(--ac)">{fmtnum(total_mag)}</div><div class="mc-sub">komada</div></div>
  <div class="mag-card"><div class="mc-label">Ø Mesečna prodaja</div><div class="mc-val" style="color:var(--t1)">{fmtnum(total_avg)}</div><div class="mc-sub">kom/mesec</div></div>
  <div class="mag-card"><div class="mc-label">Pokrivenost</div><div class="mc-val" style="color:{'var(--grn)' if total_days > 90 else 'var(--red)'}">{total_months_str} mes</div><div class="mc-sub">{total_days} dana</div></div>
  <div class="mag-card"><div class="mc-label">Artikala</div><div class="mc-val" style="color:var(--t2)">{len(df_mag)}</div><div class="mc-sub">u magacinu</div></div>
</div>'''

    grupa_cards = '<div class="mag-cards">'
    for g in sorted(mag_grupe_summary.keys()):
        gs = mag_grupe_summary[g]
        dd = round(gs['kol'] / (gs['avg'] / 30)) if gs['avg'] > 0 else 9999
        mm = f"{dd / 30:.1f}" if dd < 9999 else "∞"
        clr = 'var(--grn)' if dd > 90 else ('var(--red)' if dd <= 30 else '#92400e')
        grupa_cards += f'<div class="mag-card"><div class="mc-label">{g}</div><div class="mc-val" style="color:{clr}">{mm} mes</div><div class="mc-sub">{fmtnum(gs["kol"])} kom · ø {fmtnum(gs["avg"])}/mes</div></div>'
    grupa_cards += '</div>'

    CSS = _CSS
    JS = ('var PRODAJA_DATA=' + prodaja_json + ';\nvar PROFIT_DATA=' + profit_json +  # noqa
          ';\nvar DRV_DATA=' + drv_json + ';\nvar NUM_MONTHS=' + str(num_s) +
          ';\nvar NUM_PROFIT_MONTHS=' + str(num_p) + ';\nvar NUM_DRV_MONTHS=' + str(num_d) + ';\n' +
          _JS_BODY +
          '\nvar ANALITIKA_DATA=' + analitika_data_json + ';\n' + _JS_ANALITIKA +
          '\nvar AKCIJA_DATA=' + akcija_data_json + ';\nvar OBRT_DATA=' + obrt_chart_json + ';\n' + _JS_AKCIJA_OBRT)

    if potpun:
        tabs_block = ('<div class="tab active" onclick="showTab(\'prodaja\')">📊 PRODAJA</div>'
                      '<div class="tab" onclick="showTab(\'drv\')">🎯 USPEŠNOST AKCIJE</div>'
                      '<div class="tab" onclick="showTab(\'profit\')">💰 PROFITABILNOST</div>'
                      '<div class="tab" onclick="showTab(\'zalihe\')">📦 ZALIHE</div>')
        filters_profit_block = f'<div class="filters hidden" id="filters-profit"><label>Sistem:</label><select id="fp-sistem" onchange="applyProfitFilters()"><option value="">Svi sistemi</option>{sistem_options}</select><span style="width:12px"></span><label>Troškovi:</label>{trosak_checks}<button class="reset-btn" onclick="resetProfitFilters()">✕ Reset</button></div>'
        filters_zalihe_block = f'''<div class="filters hidden" id="filters-zalihe"><label>Sistem:</label><select id="fz-sistem" onchange="applyZaliheFilters()"><option value="">Svi sistemi</option>{sistem_options}</select><button class="reset-btn" onclick="document.getElementById('fz-sistem').value='';applyZaliheFilters()">✕ Reset</button></div>'''
        prodaja_analitika = analitika_html
        drv_akcija = akcija_analiza_html
        profit_panel_block = f'<div class="panel" id="panel-profit"><div class="tw"><table><thead>{prh}</thead><tbody id="tbody-profit">{chr(10).join(profit_rows)}</tbody></table></div><div class="ft">Klikni na sistem → detalji profita i troškova</div></div>'
        zalihe_panel_block = f'''<div class="panel" id="panel-zalihe">
<div class="mag-section"><h3>STANJE MAGACINA</h3><div class="mag-sub">{mag_info}</div>{cards_html}{grupa_cards}<div class="tw" style="padding:0"><table class="mag-table"><thead>{mh}</thead><tbody>{chr(10).join(mag_rows)}</tbody></table></div></div>
<div class="zal-divider">ZALIHE PO SISTEMIMA (poslednje stanje)</div>
<div class="tw"><table><thead>{zsh}</thead><tbody id="tbody-zalihe">{chr(10).join(zal_rows)}</tbody></table></div>
{obrt_lagera_html}</div>'''
    else:
        tabs_block = ('<div class="tab active" onclick="showTab(\'prodaja\')">📊 PRODAJA</div>'
                      '<div class="tab" onclick="showTab(\'drv\')">🎯 USPEŠNOST AKCIJE</div>')
        filters_profit_block = '';
        filters_zalihe_block = '';
        prodaja_analitika = '';
        drv_akcija = '';
        profit_panel_block = '';
        zalihe_panel_block = ''

    final = f'''<!DOCTYPE html>
<html lang="sr"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Dashboard Akcija Vape</title>
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600;700&family=Outfit:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
<style>{CSS}</style></head><body>
<div class="hdr"><div><h1>VAPE DASHBOARD</h1><div class="sub">{info}</div></div>
  <div class="hb" style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">{badge_html}
    <button onclick="toggleAll(true)" style="background:var(--acd);color:var(--ac);border:1px solid rgba(37,99,235,0.15)">Otvori sve</button>
    <button onclick="toggleAll(false)" style="background:rgba(90,95,122,0.06);color:var(--t2);border:1px solid var(--bd2)">Zatvori sve</button></div></div>
<div class="toolbar"><div class="tabs">{tabs_block}</div></div>
<div class="filters" id="filters-prodaja"><label>Sistem:</label><select id="f-sistem" onchange="applyFilters()"><option value="">Svi sistemi</option>{sistem_options}</select><label>Grupa:</label><select id="f-grupa" onchange="applyFilters()"><option value="">Sve grupe</option>{grupa_options}</select><button class="reset-btn" onclick="resetFilters()">✕ Reset</button></div>
{filters_profit_block}
<div class="filters hidden" id="filters-drv"><label>Sistem:</label><select id="fd-sistem" onchange="applyDrvFilters()"><option value="">Svi sistemi</option>{sistem_options}</select><button class="reset-btn" onclick="document.getElementById('fd-sistem').value='';applyDrvFilters()">✕ Reset</button></div>
{filters_zalihe_block}
<div class="legend" id="leg"><b>NERD:</b><span class="lc" style="background:#90EE90">1390</span><span class="lc" style="background:#FFD1DC">1300</span><span class="lc" style="background:#FFB6C1">1290</span><span class="lc" style="background:#FF69B4;color:#fff">1190</span><span class="lc" style="background:#C71585;color:#fff">990</span><span style="width:8px"></span><b>HQD:</b><span class="lc" style="background:#90EE90">890</span><span class="lc" style="background:#FFD1DC">800</span><span class="lc" style="background:#FFB6C1">790</span><span class="lc" style="background:#FF69B4;color:#fff">730</span><span class="lc" style="background:#C71585;color:#fff">690</span><span style="width:8px"></span><b>NERD 2000:</b><span class="lc" style="background:#DC2626;color:#fff">590</span></div>
<div class="panel active" id="panel-prodaja"><div class="tw"><table><thead>{ph}</thead><tbody id="tbody-prodaja">{chr(10).join(prodaja_rows)}</tbody></table></div><div class="ft">Klikni na sistem → grupe · Klikni na grupu → artikli</div>{prodaja_analitika}</div>
{profit_panel_block}
<div class="panel" id="panel-drv"><div class="tw"><table><thead>{drvh}</thead><tbody id="tbody-drv">{chr(10).join(drv_rows)}</tbody></table></div><div class="ft">Profitabilnost po odjavi · marketing + dodatni mesečni trošak</div>{drv_akcija}</div>
{zalihe_panel_block}
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<script>{JS}</script></body></html>'''

    # === EXCEL izvoz (za direktore da rade sa poljima) ===
    xlsx_bytes = None
    try:
        _xbuf = _io.BytesIO()
        with pd.ExcelWriter(_xbuf, engine='openpyxl') as _xw:
            _r1 = []
            for s in sistemi_lista:
                row = {'SISTEM': s}
                for i, nm in enumerate(nazivi): row[nm] = sistem_monthly[s][i]
                row['TOTAL'] = sum(sistem_monthly[s])
                _r1.append(row)
            pd.DataFrame(_r1).to_excel(_xw, sheet_name='Prodaja sistem-mesec', index=False)
            _r2 = []
            for s in sistemi_lista:
                for grp, vals in prodaja_data_js[s].items():
                    row = {'SISTEM': s, 'GRUPA': grp}
                    for i, nm in enumerate(nazivi): row[nm] = vals[i]
                    row['TOTAL'] = sum(vals)
                    _r2.append(row)
            pd.DataFrame(_r2).to_excel(_xw, sheet_name='Prodaja po grupama', index=False)
            _r3 = []
            for s in sistemi_lista:
                dv = drv_data_js[s]
                row = {'SISTEM': s};
                tot = 0
                for i, nm in enumerate(nazivi_drv):
                    neto = dv['profit'][i] - dv['mkt'][i] - (dv['dod'][i] if dv.get('dod') else 0)
                    row[nm] = neto;
                    tot += neto
                row['TOTAL'] = tot
                _r3.append(row)
            pd.DataFrame(_r3).to_excel(_xw, sheet_name='Neto profit sistem-mesec', index=False)
        xlsx_bytes = _xbuf.getvalue()
    except Exception:
        xlsx_bytes = None

    # Strukturirani podaci prodaje po sistemu (za direktorski „Detaljan izveštaj po sistemima")
    sistem_prodaja = {
        "nazivi": nazivi,
        "mesec_label": poslednji_naziv,
        "po_sistemu": {s: {"total": sistem_monthly.get(s, []), "grupe": prodaja_data_js.get(s, {})}
                       for s in sistemi_lista},
    }

    return final, xlsx_bytes, poslednji_naziv, sistem_prodaja


# ------- veliki statični delovi (CSS i JS) -------
_CSS = '''*{margin:0;padding:0;box-sizing:border-box}
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
.tcb{display:inline-flex;align-items:center;gap:4px;font-family:'IBM Plex Mono',monospace;font-size:10px;color:var(--t2);cursor:pointer;padding:3px 8px;border-radius:6px;background:var(--bg2);border:1px solid var(--bd);transition:all .15s;text-transform:none;letter-spacing:0}
.tcb:hover{border-color:var(--ac)}.tcb input{accent-color:var(--ac);cursor:pointer}
.tcb input:not(:checked)+span{color:var(--t3);text-decoration:line-through}
.legend{display:flex;gap:8px;align-items:center;flex-wrap:wrap;padding:8px 28px;font-size:10px;color:var(--t3);background:var(--bg);border-bottom:1px solid var(--bd)}
.legend b{font-family:'IBM Plex Mono',monospace;font-weight:700}
.lc{display:inline-block;padding:2px 7px;border-radius:4px;font-weight:700;font-size:9px;font-family:'IBM Plex Mono',monospace;color:#1a1a2e;box-shadow:0 1px 2px rgba(0,0,0,0.1)}
.panel{display:none}.panel.active{display:block}
.tw{overflow-x:auto;padding:4px 16px 40px}
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
.an{padding-left:36px;color:var(--t3);font-size:10px}.an::before{content:'↳ ';color:var(--bd2)}
.n{text-align:right}.nb{text-align:right;font-weight:600}.nt{text-align:right;font-weight:800;color:var(--ac)}
.pc{color:var(--grn)}.pf{color:var(--red)}.cc{color:var(--red);font-style:italic;opacity:.75}
.cl{color:var(--red);font-style:italic;font-size:10px;opacity:.75}
.ctl{color:#b91c1c;font-weight:800;font-size:10px}.ctc{color:#b91c1c;font-weight:700}
.nl{color:var(--grn);font-weight:800;font-size:10px}.np{color:var(--grn);font-weight:700}.nn{color:var(--red);font-weight:700}
.nf{color:var(--red);font-weight:600;font-size:10px}.no{color:var(--ac);font-weight:600;font-size:10px}
.hidden{display:none}
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

_JS_BODY = r'''function showTab(n){document.querySelectorAll('.panel').forEach(function(p){p.classList.remove('active')});document.querySelectorAll('.tab').forEach(function(t){t.classList.remove('active')});var pnl=document.getElementById('panel-'+n);if(pnl)pnl.classList.add('active');document.querySelectorAll('.tab').forEach(function(t){var oc=t.getAttribute('onclick');if(oc&&oc.indexOf("'"+n+"'")>-1)t.classList.add('active')});function _sd(id,v){var e=document.getElementById(id);if(e)e.style.display=v}_sd('leg',(n==='prodaja')?'flex':'none');_sd('filters-prodaja',(n==='prodaja')?'flex':'none');_sd('filters-profit',(n==='profit')?'flex':'none');_sd('filters-drv',(n==='drv')?'flex':'none');_sd('filters-zalihe',(n==='zalihe')?'flex':'none')}
function tog(id){var btn=document.getElementById('b-'+id);if(!btn)return;var isO=btn.textContent.trim()==='−';var rows=document.querySelectorAll('tr[data-p="'+id+'"]');if(isO){rows.forEach(function(r){r.classList.add('hidden');var cb=r.querySelector('.beg');if(cb){cb.textContent='+';var cid=cb.id.replace('b-','');document.querySelectorAll('tr[data-p="'+cid+'"]').forEach(function(cr){cr.classList.add('hidden')})}});btn.textContent='+'}else{var fG=document.getElementById('f-grupa')?document.getElementById('f-grupa').value:'';rows.forEach(function(r){if(fG&&r.getAttribute('data-grupa')&&r.getAttribute('data-grupa')!==fG)return;r.classList.remove('hidden')});btn.textContent='−'}}
function toggleAll(o){var ap=document.querySelector('.panel.active');if(!ap)return;ap.querySelectorAll('.be').forEach(function(btn){var id=btn.id.replace('b-','');var rows=document.querySelectorAll('tr[data-p="'+id+'"]');rows.forEach(function(r){o?r.classList.remove('hidden'):r.classList.add('hidden')});btn.textContent=o?'−':'+'})}
function applyFilters(){var fS=document.getElementById('f-sistem').value;var fG=document.getElementById('f-grupa').value;var tbody=document.getElementById('tbody-prodaja');tbody.querySelectorAll('.be').forEach(function(btn){btn.textContent='+'});tbody.querySelectorAll('tr').forEach(function(r){var rS=r.getAttribute('data-sistem');var isSr=r.classList.contains('sr');var isGr=r.classList.contains('gr');var isAr=r.classList.contains('ar');var isSep=r.classList.contains('sep');var isTotal=r.classList.contains('totalrow');if(isTotal){r.classList.remove('hidden');return}if(fS&&rS&&rS!==fS){r.classList.add('hidden');return}if(isSr){if(fG){var hG=PRODAJA_DATA[rS]&&PRODAJA_DATA[rS][fG];if(!hG){r.classList.add('hidden');return}var cells=r.querySelectorAll('td');var gt=0;for(var i=0;i<NUM_MONTHS;i++){cells[i+1].textContent=hG[i].toLocaleString('sr-RS');gt+=hG[i]}cells[NUM_MONTHS+1].textContent=gt.toLocaleString('sr-RS')}else{var allG=PRODAJA_DATA[rS];if(allG){var cells=r.querySelectorAll('td');var sums=new Array(NUM_MONTHS).fill(0);for(var g in allG)for(var i=0;i<NUM_MONTHS;i++)sums[i]+=allG[g][i];var gt=0;for(var i=0;i<NUM_MONTHS;i++){cells[i+1].textContent=sums[i].toLocaleString('sr-RS');gt+=sums[i]}cells[NUM_MONTHS+1].textContent=gt.toLocaleString('sr-RS')}}r.classList.remove('hidden');return}if(isGr||isAr){r.classList.add('hidden');return}if(isSep){if(fS&&rS!==fS){r.classList.add('hidden');return}if(fG){var hSG=PRODAJA_DATA[rS]&&PRODAJA_DATA[rS][fG];if(!hSG){r.classList.add('hidden');return}}r.classList.remove('hidden');return}});recalcTotals(fS,fG)}
function recalcTotals(fS,fG){var t=new Array(NUM_MONTHS).fill(0);for(var s in PRODAJA_DATA){if(fS&&s!==fS)continue;for(var g in PRODAJA_DATA[s]){if(fG&&g!==fG)continue;var v=PRODAJA_DATA[s][g];for(var i=0;i<NUM_MONTHS;i++)t[i]+=v[i]}}var gt=t.reduce(function(a,b){return a+b},0);var tr=document.getElementById('prodaja-total');if(!tr)return;var c=tr.querySelectorAll('td');for(var i=1;i<=NUM_MONTHS;i++)c[i].textContent=t[i-1].toLocaleString('sr-RS');c[NUM_MONTHS+1].textContent=gt.toLocaleString('sr-RS')}
function resetFilters(){document.getElementById('f-sistem').value='';document.getElementById('f-grupa').value='';applyFilters()}
function fmtN(v){return v.toLocaleString('sr-RS')}
function applyProfitFilters(){var fS=document.getElementById('fp-sistem').value;var tbody=document.getElementById('tbody-profit');var checks=document.querySelectorAll('#filters-profit input[type=checkbox]');var ac={};checks.forEach(function(cb){ac[cb.value]=cb.checked});tbody.querySelectorAll('.be').forEach(function(btn){btn.textContent='+'});tbody.querySelectorAll('tr').forEach(function(r){var rS=r.getAttribute('data-sistem');var isSr=r.classList.contains('sr');var isSep=r.classList.contains('sep');var isTotal=r.classList.contains('totalrow');if(isTotal){r.classList.remove('hidden');return}if(isSr||isSep){if(fS&&rS&&rS!==fS)r.classList.add('hidden');else r.classList.remove('hidden');return}r.classList.add('hidden')});tbody.querySelectorAll('tr[data-trosak]').forEach(function(r){var tid=r.getAttribute('data-trosak');r.classList.toggle('excluded',!ac[tid])});var grandT=new Array(NUM_PROFIT_MONTHS).fill(0);tbody.querySelectorAll('.ukupni-row').forEach(function(ukRow){var sid=ukRow.getAttribute('data-sid');var sistem=ukRow.getAttribute('data-sistem');var pd=PROFIT_DATA[sistem];if(!pd)return;var uk=new Array(NUM_PROFIT_MONTHS).fill(0);if(ac['mkt'])for(var i=0;i<NUM_PROFIT_MONTHS;i++)uk[i]+=pd['mkt'][i];if(ac['dod']&&pd['dod'])for(var i=0;i<NUM_PROFIT_MONTHS;i++)uk[i]+=pd['dod'][i];for(var tid in pd){if(tid==='profit'||tid==='mkt'||tid==='dod')continue;if(ac[tid])for(var i=0;i<NUM_PROFIT_MONTHS;i++)uk[i]+=pd[tid][i]}var ukT=uk.reduce(function(a,b){return a+b},0);var cells=ukRow.querySelectorAll('td');for(var i=3;i<3+NUM_PROFIT_MONTHS;i++)cells[i].textContent=fmtN(uk[i-3]);cells[3+NUM_PROFIT_MONTHS].textContent=fmtN(ukT);var nr=tbody.querySelector('.neto-row[data-sid="'+sid+'"]');if(!nr)return;var p=pd['profit'];var nc=nr.querySelectorAll('td');var nt=0;for(var i=0;i<NUM_PROFIT_MONTHS;i++){var nv=p[i]-uk[i];nt+=nv;nc[3+i].textContent=fmtN(nv);nc[3+i].className='n '+(nv>=0?'np':'nn')}nc[3+NUM_PROFIT_MONTHS].textContent=fmtN(nt);nc[3+NUM_PROFIT_MONTHS].className='n '+(nt>=0?'np':'nn');nc[3+NUM_PROFIT_MONTHS].style.fontSize='13px';var sr=tbody.querySelector('.profit-sr[data-sid="'+sid+'"]');if(sr){var sc=sr.querySelectorAll('td');for(var i=3;i<3+NUM_PROFIT_MONTHS;i++){var nv2=p[i-3]-uk[i-3];sc[i].textContent=fmtN(nv2);sc[i].className='nb '+(nv2>=0?'np':'nn')}var stot=p.reduce(function(a,b){return a+b},0)-ukT;sc[3+NUM_PROFIT_MONTHS].textContent=fmtN(stot);sc[3+NUM_PROFIT_MONTHS].className='nb '+(stot>=0?'np':'nn');sc[3+NUM_PROFIT_MONTHS].style.fontSize='13px'}if(!fS||sistem===fS){for(var i=0;i<NUM_PROFIT_MONTHS;i++)grandT[i]+=p[i]-uk[i]}});var tRow=document.getElementById('profit-total');if(tRow){var tc=tRow.querySelectorAll('td');var ggt=grandT.reduce(function(a,b){return a+b},0);for(var i=3;i<3+NUM_PROFIT_MONTHS;i++){tc[i].textContent=fmtN(grandT[i-3]);tc[i].className='nb total-cell '+(grandT[i-3]>=0?'np':'nn')}tc[3+NUM_PROFIT_MONTHS].textContent=fmtN(ggt);tc[3+NUM_PROFIT_MONTHS].className='nb total-cell '+(ggt>=0?'np':'nn');tc[3+NUM_PROFIT_MONTHS].style.fontSize='13px'}}
function resetProfitFilters(){document.getElementById('fp-sistem').value='';document.querySelectorAll('#filters-profit input[type=checkbox]').forEach(function(cb){cb.checked=true});applyProfitFilters()}
function applyDrvFilters(){var fS=document.getElementById('fd-sistem').value;var tbody=document.getElementById('tbody-drv');tbody.querySelectorAll('.be').forEach(function(btn){btn.textContent='+'});tbody.querySelectorAll('tr').forEach(function(r){var rS=r.getAttribute('data-sistem');var isSr=r.classList.contains('sr');var isSep=r.classList.contains('sep');var isTotal=r.classList.contains('totalrow');if(isTotal){r.classList.remove('hidden');return}if(isSr||isSep){if(fS&&rS&&rS!==fS)r.classList.add('hidden');else r.classList.remove('hidden');return}r.classList.add('hidden')});var grandT=new Array(NUM_DRV_MONTHS).fill(0);for(var s in DRV_DATA){if(fS&&s!==fS)continue;var d=DRV_DATA[s];for(var i=0;i<NUM_DRV_MONTHS;i++){var dod=(d['dod']&&d['dod'][i])?d['dod'][i]:0;grandT[i]+=d['profit'][i]-d['mkt'][i]-dod}}var tRow=document.getElementById('drv-total');if(tRow){var tc=tRow.querySelectorAll('td');var ggt=grandT.reduce(function(a,b){return a+b},0);for(var i=3;i<3+NUM_DRV_MONTHS;i++){tc[i].textContent=fmtN(grandT[i-3]);tc[i].className='nb total-cell '+(grandT[i-3]>=0?'np':'nn')}tc[3+NUM_DRV_MONTHS].textContent=fmtN(ggt);tc[3+NUM_DRV_MONTHS].className='nb total-cell '+(ggt>=0?'np':'nn');tc[3+NUM_DRV_MONTHS].style.fontSize='13px'}}
function applyZaliheFilters(){var fS=document.getElementById('fz-sistem').value;var tbody=document.getElementById('tbody-zalihe');tbody.querySelectorAll('.be').forEach(function(btn){btn.textContent='+'});tbody.querySelectorAll('tr').forEach(function(r){var rS=r.getAttribute('data-sistem');var isSr=r.classList.contains('sr');var isSep=r.classList.contains('sep');var isTotal=r.classList.contains('totalrow');if(isTotal){r.classList.remove('hidden');return}if(isSr||isSep){if(fS&&rS&&rS!==fS)r.classList.add('hidden');else r.classList.remove('hidden');return}r.classList.add('hidden')})}'''

_JS_ANALITIKA = r'''var chartsInitialized=false;
function initAnalitikaCharts(){
  if(chartsInitialized||typeof Chart==='undefined')return;
  chartsInitialized=true;
  var d=ANALITIKA_DATA;
  var monoFont={family:"IBM Plex Mono",size:9};
  var monoFont10={family:"IBM Plex Mono",size:10};
  var commonTooltip={backgroundColor:"#fff",titleColor:"#1a1a2e",bodyColor:"#5a5f7a",borderColor:"#e2e6ef",borderWidth:1,padding:8,titleFont:monoFont10,bodyFont:monoFont10};
  function makeMini(id,obj){if(!obj)return;var el=document.getElementById(id);if(!el)return;new Chart(el,{type:"bar",data:{labels:obj.labels,datasets:[{data:obj.values,backgroundColor:["#cbd5e0",obj.color],borderRadius:4}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{enabled:false}},scales:{y:{display:false,beginAtZero:false},x:{display:false}}}});}
  makeMini("mini1",d.mini1);makeMini("mini2",d.mini2);makeMini("mini3",d.mini3);
  var bigEl=document.getElementById("bigTrendChart");
  if(bigEl){var bg=d.big;var n=bg.values.length;var rolling6=bg.values.map(function(_,i,a){if(i<5)return null;var s=0;for(var k=i-5;k<=i;k++)s+=a[k];return Math.round(s/6);});var yoyArr=new Array(n).fill(null);for(var i=0;i<n;i++){var prevI=i-12;if(prevI>=0)yoyArr[i]=bg.values[prevI];}
    new Chart(bigEl,{type:"line",data:{labels:bg.labels,datasets:[{label:"Mesečna prodaja",data:bg.values,borderColor:"#2563eb",backgroundColor:"rgba(37,99,235,0.08)",fill:true,tension:0.35,borderWidth:2.5,pointRadius:3.5,pointBackgroundColor:"#2563eb",pointBorderColor:"#fff",pointBorderWidth:1.5},{label:"6M prosek",data:rolling6,borderColor:"#dc2626",borderDash:[6,4],borderWidth:1.8,pointRadius:0,fill:false,tension:0.35},{label:"YoY",data:yoyArr,borderColor:"#16a34a",borderWidth:0,pointRadius:3,pointBackgroundColor:"#16a34a",pointBorderColor:"#fff",pointBorderWidth:1,showLine:false}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:Object.assign({mode:"index",intersect:false},commonTooltip)},scales:{y:{beginAtZero:false,grid:{color:"rgba(0,0,0,0.04)"},ticks:{font:monoFont10,color:"#8b90a5"}},x:{grid:{display:false},ticks:{font:monoFont,color:"#8b90a5",maxRotation:45}}}}});}
  Object.keys(d.grupe).forEach(function(grKey){var g=d.grupe[grKey];var el=document.getElementById("chart-"+grKey);if(!el)return;var labels=g.has_yoy?g.labels:g.labels.slice(0,3);var values=g.has_yoy?g.values:g.values.slice(0,3);var colors=values.map(function(_,i){return i===0?g.main_color:"#cbd5e0";});new Chart(el,{type:"bar",data:{labels:labels,datasets:[{data:values,backgroundColor:colors,borderRadius:6,borderSkipped:false}]},options:{indexAxis:"y",responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:commonTooltip},scales:{x:{beginAtZero:false,grid:{color:"rgba(0,0,0,0.04)"},ticks:{font:monoFont10,color:"#8b90a5",callback:function(v){return v.toLocaleString("sr-RS");}}},y:{grid:{display:false},ticks:{font:{family:"IBM Plex Mono",size:10,weight:"bold"},color:"#1a1a2e"}}}}});});
}'''

_JS_AKCIJA_OBRT = r'''var akcijaChartsInit=false;
function initAkcijaCharts(){
  if(akcijaChartsInit||typeof Chart==='undefined')return;akcijaChartsInit=true;
  var monoFont={family:"IBM Plex Mono",size:9};var monoFont10={family:"IBM Plex Mono",size:10};
  var commonTooltip={backgroundColor:"#fff",titleColor:"#1a1a2e",bodyColor:"#5a5f7a",borderColor:"#e2e6ef",borderWidth:1,padding:8,titleFont:monoFont10,bodyFont:monoFont10};
  Object.keys(AKCIJA_DATA).forEach(function(safeId){var d=AKCIJA_DATA[safeId];var el=document.getElementById("ak-chart-"+safeId);if(!el)return;var colors=d.values.map(function(_,i){return d.is_akcija[i]?"#0d9488":"#cbd5e1";});
    new Chart(el,{data:{labels:d.labels,datasets:[{type:"bar",label:"Količina",data:d.values,backgroundColor:colors,borderRadius:3,yAxisID:"y",order:2},{type:"line",label:"Bruto profit",data:d.profit,borderColor:"#1e3a8a",backgroundColor:"#1e3a8a",borderWidth:2.5,pointRadius:3,pointBackgroundColor:"#1e3a8a",pointBorderColor:"#fff",pointBorderWidth:1,tension:0.3,fill:false,yAxisID:"y1",order:1}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:Object.assign({mode:"index",intersect:false},commonTooltip,{callbacks:{label:function(c){if(c.dataset.type==="line")return "Bruto profit: "+Math.round(c.parsed.y).toLocaleString("sr-RS")+" RSD";return (d.is_akcija[c.dataIndex]?"AKCIJA: ":"redovno: ")+c.parsed.y.toLocaleString("sr-RS")+" kom";}}})},scales:{y:{position:"left",beginAtZero:true,grid:{color:"rgba(0,0,0,0.04)"},ticks:{font:monoFont,color:"#8b90a5"}},y1:{position:"right",beginAtZero:true,grid:{display:false},ticks:{font:monoFont,color:"#1e3a8a",callback:function(v){return (v/1000).toFixed(0)+"k";}}},x:{grid:{display:false},ticks:{font:{family:"IBM Plex Mono",size:8},color:"#8b90a5",maxRotation:45}}}}});});
}
var obrtChartInit=false;
function initObrtChart(){
  if(obrtChartInit||typeof Chart==='undefined')return;var el=document.getElementById("obrtChart");if(!el||!OBRT_DATA||!OBRT_DATA.sistemi||OBRT_DATA.sistemi.length===0)return;obrtChartInit=true;
  function colorFor(v){if(v<=3)return "#16a34a";if(v<=6)return "#0d9488";if(v<=12)return "#f59e0b";return "#dc2626";}
  var colors=OBRT_DATA.meseci.map(colorFor);
  new Chart(el,{type:"bar",data:{labels:OBRT_DATA.sistemi,datasets:[{data:OBRT_DATA.meseci,backgroundColor:colors,borderRadius:4}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{backgroundColor:"#fff",titleColor:"#1a1a2e",bodyColor:"#5a5f7a",borderColor:"#e2e6ef",borderWidth:1,padding:10,titleFont:{family:"IBM Plex Mono",size:11},bodyFont:{family:"IBM Plex Mono",size:10},callbacks:{label:function(c){var i=c.dataIndex;return [c.parsed.y.toFixed(1)+" meseci za obrt","Stanje "+OBRT_DATA.last_mes[i]+": "+OBRT_DATA.lager[i].toLocaleString("sr-RS")+" kom","Ø prodaja: "+OBRT_DATA.prodaja_avg[i].toLocaleString("sr-RS")+"/mes"];}}}},scales:{y:{beginAtZero:true,grid:{color:"rgba(0,0,0,0.04)"},ticks:{font:{family:"IBM Plex Mono",size:10},color:"#8b90a5",callback:function(v){return v+" mes";}}},x:{grid:{display:false},ticks:{font:{family:"IBM Plex Mono",size:10,weight:"bold"},color:"#1a1a2e",maxRotation:45,minRotation:45}}}}});
}
var _origShowTab=showTab;
showTab=function(n){_origShowTab(n);if(n==="prodaja")setTimeout(initAnalitikaCharts,50);if(n==="drv")setTimeout(initAkcijaCharts,50);if(n==="zalihe")setTimeout(initObrtChart,50);};
if(document.readyState!=="loading"){setTimeout(initAnalitikaCharts,100);}else{document.addEventListener("DOMContentLoaded",function(){setTimeout(initAnalitikaCharts,100);});}'''
