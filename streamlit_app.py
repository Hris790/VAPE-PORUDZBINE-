"""
💨 VAPE PORUDŽBINE — Streamlit Web Aplikacija
===============================================
Pokretanje lokalno:   streamlit run streamlit_app.py
Hosting:              Streamlit Community Cloud (besplatno)

Potrebno:  pip install streamlit pandas openpyxl numpy
"""

import io
import datetime
import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import streamlit as st

# ──────────────────────────── PAGE CONFIG ────────────────────────────

st.set_page_config(
    page_title="VAPE Porudžbine",
    page_icon="💨",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ──────────────────────────── CUSTOM CSS ────────────────────────────

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap');

    /* Main background */
    .stApp {
        background: linear-gradient(160deg, #fdf2f8 0%, #f5f0ff 40%, #eff6ff 100%);
        font-family: 'Poppins', sans-serif;
    }

    /* Sidebar */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #7c3aed 0%, #a855f7 50%, #c084fc 100%) !important;
    }
    section[data-testid="stSidebar"] * {
        color: white !important;
    }
    section[data-testid="stSidebar"] .stTextInput > div > div > input,
    section[data-testid="stSidebar"] .stNumberInput > div > div > input,
    section[data-testid="stSidebar"] .stTextArea > div > div > textarea {
        background: rgba(255,255,255,0.15) !important;
        border: 1px solid rgba(255,255,255,0.3) !important;
        color: white !important;
        border-radius: 8px !important;
    }
    section[data-testid="stSidebar"] label {
        color: rgba(255,255,255,0.9) !important;
        font-weight: 500 !important;
    }

    /* Headers */
    h1, h2, h3 {
        font-family: 'Poppins', sans-serif !important;
    }

    /* Cards */
    .metric-card {
        background: white;
        border-radius: 16px;
        padding: 20px 24px;
        box-shadow: 0 2px 12px rgba(124, 58, 237, 0.08);
        border: 1px solid rgba(124, 58, 237, 0.1);
        text-align: center;
    }
    .metric-value {
        font-size: 32px;
        font-weight: 700;
        background: linear-gradient(135deg, #7c3aed, #ec4899);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    .metric-label {
        font-size: 13px;
        color: #888;
        margin-top: 4px;
    }

    /* Buttons */
    .stButton > button {
        background: linear-gradient(135deg, #a855f7 0%, #ec4899 100%) !important;
        color: white !important;
        border: none !important;
        border-radius: 12px !important;
        padding: 12px 32px !important;
        font-weight: 600 !important;
        font-size: 16px !important;
        letter-spacing: 0.5px;
        transition: all 0.3s ease !important;
        box-shadow: 0 4px 15px rgba(168, 85, 247, 0.3) !important;
    }
    .stButton > button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 20px rgba(168, 85, 247, 0.4) !important;
    }

    /* Download button */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important;
        color: white !important;
        border: none !important;
        border-radius: 12px !important;
        padding: 12px 32px !important;
        font-weight: 600 !important;
        box-shadow: 0 4px 15px rgba(16, 185, 129, 0.3) !important;
    }

    /* File uploader */
    .stFileUploader > div {
        border-radius: 16px !important;
        border: 2px dashed rgba(168, 85, 247, 0.3) !important;
        background: rgba(168, 85, 247, 0.03) !important;
    }

    /* Progress bar */
    .stProgress > div > div {
        background: linear-gradient(90deg, #a855f7, #ec4899) !important;
    }

    /* Expander */
    .streamlit-expanderHeader {
        font-weight: 600 !important;
        color: #7c3aed !important;
    }

    /* Tabs */
    .stTabs [data-baseweb="tab"] {
        font-weight: 500;
    }
    .stTabs [aria-selected="true"] {
        color: #7c3aed !important;
    }

    /* Tables */
    .stDataFrame {
        border-radius: 12px !important;
        overflow: hidden;
    }

    /* Success/info boxes */
    .success-box {
        background: linear-gradient(135deg, rgba(16,185,129,0.1), rgba(5,150,105,0.05));
        border: 1px solid rgba(16,185,129,0.2);
        border-radius: 12px;
        padding: 16px 20px;
    }

    /* Header banner */
    .header-banner {
        background: linear-gradient(135deg, #7c3aed 0%, #a855f7 30%, #ec4899 70%, #f472b6 100%);
        border-radius: 16px;
        padding: 24px 32px;
        color: white;
        margin-bottom: 24px;
        box-shadow: 0 4px 20px rgba(124, 58, 237, 0.25);
    }
    .header-title {
        font-size: 28px;
        font-weight: 700;
        margin: 0;
        letter-spacing: -0.5px;
    }
    .header-sub {
        font-size: 14px;
        opacity: 0.85;
        margin-top: 4px;
    }

    /* Log area */
    .log-area {
        background: #1a1a2e;
        border-radius: 12px;
        padding: 16px;
        color: #a0a0b0;
        font-family: 'Consolas', monospace;
        font-size: 12px;
        max-height: 300px;
        overflow-y: auto;
    }
    .log-success { color: #00d4aa; }
    .log-step { color: #b4d7e8; }
</style>
""", unsafe_allow_html=True)

# ──────────────────────────── KONSTANTE ────────────────────────────

DEFAULT_EXCLUDED = "1023, 1027, 1034, 1043, 1057, 1060, 1061, 1076, 1315, 1347, 1349, 1359"
WMA_WEIGHTS = np.array([0.05, 0.10, 0.15, 0.25, 0.45])


# ──────────────────────────── ENGINE ────────────────────────────

class PredictionEngine:
    def __init__(self, file_bytes, excluded_ids, alpha, beta, min_lager, min_order):
        self.file_bytes = file_bytes
        self.excluded = excluded_ids
        self.alpha = alpha
        self.beta = beta
        self.min_lager = min_lager
        self.min_order = min_order
        self.logs = []
        self.adjustments = []

    def log(self, msg):
        self.logs.append(msg)

    def run(self, progress_bar, status_text):
        progress_bar.progress(5, "Učitavanje sheetova...")
        self._load_sheets()

        progress_bar.progress(15, "Priprema podataka...")
        self._prepare_lookups()

        progress_bar.progress(25, "Računanje povrata i korekcija...")
        self._compute_povrat_korekcija()

        progress_bar.progress(40, "Mesečni pregled...")
        self._build_monthly()

        progress_bar.progress(55, "OOS-korigovana predikcija...")
        self._predict_all()

        progress_bar.progress(70, "Spajanje sa lagerom...")
        self._merge_lager()

        progress_bar.progress(80, "Računanje porudžbina...")
        self._compute_orders()

        progress_bar.progress(85, f"Pravilo min {self.min_order} po objektu...")
        self._apply_min_order()

        progress_bar.progress(100, "Gotovo!")
        status_text.empty()
        return self.df_result

    def _load_sheets(self):
        xls = pd.ExcelFile(io.BytesIO(self.file_bytes))
        sheet_map = {s.strip().lower(): s for s in xls.sheet_names}

        def find(keywords):
            for kw in keywords:
                for nl, no in sheet_map.items():
                    if kw in nl:
                        return no
            return None

        s_prod = find(['prodaja'])
        s_start = find(['startni'])
        s_pov = find(['povrat'])
        s_tl = find(['trenutni'])

        if not s_prod: raise ValueError("Nema sheeta 'prodaja'!")
        if not s_start: raise ValueError("Nema sheeta 'startni lager'!")

        self.prodaja = pd.read_excel(xls, sheet_name=s_prod)
        self.prodaja.columns = [c.strip() for c in self.prodaja.columns]
        self.log(f"✅ Prodaja: {len(self.prodaja)} redova")

        self.startni = pd.read_excel(xls, sheet_name=s_start)
        self.startni.columns = [c.strip() for c in self.startni.columns]
        self.log(f"✅ Startni lager: {len(self.startni)} redova")

        if s_pov:
            self.povrat_df = pd.read_excel(xls, sheet_name=s_pov)
            self.povrat_df.columns = [c.strip() for c in self.povrat_df.columns]
            self.log(f"✅ Povrat: {len(self.povrat_df)} redova")
        else:
            self.povrat_df = pd.DataFrame()
            self.log("⚠️ Nema sheeta 'povrat'")

        if s_tl:
            self.trenutni = pd.read_excel(xls, sheet_name=s_tl)
            self.trenutni.columns = [c.strip() for c in self.trenutni.columns]
            self.log(f"✅ Trenutni lager: {len(self.trenutni)} redova")
        else:
            self.trenutni = pd.DataFrame()
            self.log("⚠️ Nema sheeta 'trenutni lager'")

        self.meseci_order = sorted(
            self.prodaja[['Godina', 'Mesec']].drop_duplicates().values.tolist())
        mn = {1:'Jan',2:'Feb',3:'Mar',4:'Apr',5:'Maj',6:'Jun',
              7:'Jul',8:'Avg',9:'Sep',10:'Okt',11:'Nov',12:'Dec'}
        self.mesec_labels = [f"{mn.get(int(m),'?')} {int(g)}" for g, m in self.meseci_order]
        self.log(f"📅 Meseci: {', '.join(self.mesec_labels)}")

    def _prepare_lookups(self):
        kp = self.prodaja[['ID KOMITENTA','id artikla','Naziv artikla','Grupa']].drop_duplicates()
        ks = self.startni[['ID KOMITENTA','id artikla','Naziv artikla','Grupa']].drop_duplicates()
        self.all_keys = pd.concat([kp, ks]).drop_duplicates().sort_values(
            ['ID KOMITENTA','id artikla']).reset_index(drop=True)

        self.startni_dict = {(r['ID KOMITENTA'], r['id artikla']): r['Kolicina']
                             for _, r in self.startni.iterrows()}

        self.has_promet = 'PROMET KA NJIMA' in self.prodaja.columns
        self.prodaja_dict = {}
        for _, r in self.prodaja.iterrows():
            key = (r['ID KOMITENTA'], r['id artikla'], r['Godina'], r['Mesec'])
            prom = r['PROMET KA NJIMA'] if self.has_promet else 0
            self.prodaja_dict[key] = (
                r.get('Prodata Kolicina', r.get('Kolicina', 0)),
                r.get('Lager', 0),
                prom if not pd.isna(prom) else 0)

        self.povrat_total = {}
        if len(self.povrat_df) > 0:
            ic = [c for c in self.povrat_df.columns if 'id' in c.lower() and 'artikl' in c.lower()]
            mc = [c for c in self.povrat_df.columns if 'mesec' in c.lower()]
            gc = [c for c in self.povrat_df.columns if 'godin' in c.lower()]
            kc = [c for c in self.povrat_df.columns if 'količ' in c.lower() or 'kolicin' in c.lower()]
            if ic and mc and gc and kc:
                for _, r in self.povrat_df.iterrows():
                    key = (r[ic[0]], r[gc[0]], r[mc[0]])
                    self.povrat_total[key] = self.povrat_total.get(key, 0) + r[kc[0]]

        self.trenutni_dict = {}
        if len(self.trenutni) > 0:
            ikc = [c for c in self.trenutni.columns if 'komitent' in c.lower()]
            iac = [c for c in self.trenutni.columns if 'artikl' in c.lower() and 'id' in c.lower()]
            lc = [c for c in self.trenutni.columns if 'lager' in c.lower()]
            if ikc and iac and lc:
                for _, r in self.trenutni.iterrows():
                    k, a = r[ikc[0]], r[iac[0]]
                    if pd.notna(k) and pd.notna(a):
                        self.trenutni_dict[(int(k), int(a))] = int(r[lc[0]]) if pd.notna(r[lc[0]]) else 0

        self.log(f"📊 Kombinacija: {len(self.all_keys)}")

    def _compute_povrat_korekcija(self):
        self.final_povrat = {}
        self.final_korekcija = {}
        if not self.has_promet or not self.povrat_total:
            return

        implied = {}
        for _, k in self.all_keys.iterrows():
            idk, ida = k['ID KOMITENTA'], k['id artikla']
            poc = self.startni_dict.get((idk, ida), 0)
            for god, mes in self.meseci_order:
                pv, lv, tv = self.prodaja_dict.get((idk, ida, god, mes), (0, 0, 0))
                lv = lv if not pd.isna(lv) else 0
                implied[(idk, ida, god, mes)] = poc + tv - pv - lv
                poc = lv

        all_art = set(list(self.prodaja['id artikla'].unique()) +
                      list(self.startni['id artikla'].unique()))
        for god, mes in self.meseci_order:
            for ida in all_art:
                ap = self.povrat_total.get((ida, god, mes), 0)
                pi = {k['ID KOMITENTA']: implied.get((k['ID KOMITENTA'], ida, god, mes), 0)
                      for _, k in self.all_keys[self.all_keys['id artikla']==ida].iterrows()
                      if implied.get((k['ID KOMITENTA'], ida, god, mes), 0) > 0}
                ni = {k['ID KOMITENTA']: implied.get((k['ID KOMITENTA'], ida, god, mes), 0)
                      for _, k in self.all_keys[self.all_keys['id artikla']==ida].iterrows()
                      if implied.get((k['ID KOMITENTA'], ida, god, mes), 0) < 0}
                tp = sum(pi.values())
                if ap > 0 and tp > 0:
                    raw = {i: ap * (v / tp) for i, v in pi.items()}
                    fl = {i: int(v) for i, v in raw.items()}
                    d = ap - sum(fl.values())
                    rem = {i: raw[i]-fl[i] for i in raw}
                    for j, i in enumerate(sorted(rem, key=rem.get, reverse=True)):
                        if j < int(d): fl[i] += 1
                    for i, pval in fl.items():
                        self.final_povrat[(i, ida, god, mes)] = pval
                        self.final_korekcija[(i, ida, god, mes)] = pi[i] - pval
                elif ap == 0:
                    for i, v in pi.items():
                        self.final_korekcija[(i, ida, god, mes)] = v
                for i, v in ni.items():
                    self.final_korekcija[(i, ida, god, mes)] = \
                        self.final_korekcija.get((i, ida, god, mes), 0) + v

    def _build_monthly(self):
        rows = []
        for _, k in self.all_keys.iterrows():
            idk, ida = k['ID KOMITENTA'], k['id artikla']
            poc = self.startni_dict.get((idk, ida), 0)
            row = {'ID KOMITENTA': idk, 'id artikla': ida,
                   'Naziv artikla': k['Naziv artikla'], 'Grupa': k['Grupa']}
            for i, (god, mes) in enumerate(self.meseci_order):
                lb = self.mesec_labels[i]
                pv, lv, tv = self.prodaja_dict.get((idk, ida, god, mes), (0, 0, 0))
                lv = lv if not pd.isna(lv) else 0
                row[f'{lb}_Pocetno'] = poc
                if self.has_promet:
                    row[f'{lb}_Promet'] = tv if not pd.isna(tv) else 0
                    row[f'{lb}_Prodaja'] = pv
                    row[f'{lb}_Povrat'] = self.final_povrat.get((idk, ida, god, mes), 0)
                    row[f'{lb}_Korekcija'] = self.final_korekcija.get((idk, ida, god, mes), 0)
                else:
                    row[f'{lb}_Prodaja'] = pv
                    row[f'{lb}_UlazPovrat'] = lv - poc + pv
                poc = lv
            rows.append(row)
        self.df_monthly = pd.DataFrame(rows)

    def _predict_all(self):
        analysis = []
        for _, k in self.all_keys.iterrows():
            idk, ida = k['ID KOMITENTA'], k['id artikla']
            poc = self.startni_dict.get((idk, ida), 0)
            sales, oos, pocs = [], [], []
            for god, mes in self.meseci_order:
                pv, lv, _ = self.prodaja_dict.get((idk, ida, god, mes), (0, 0, 0))
                lv = lv if not pd.isna(lv) else 0
                sales.append(pv); oos.append(1 if poc == 0 else 0); pocs.append(poc)
                poc = lv
            analysis.append({'idk': idk, 'ida': ida,
                             'sales': np.array(sales, dtype=float),
                             'oos': np.array(oos), 'poc': np.array(pocs, dtype=float),
                             'lager_last': lv})

        preds = {}
        for it in analysis:
            s, o, p = it['sales'], it['oos'], it['poc']
            n = len(s)
            noos = s[o == 0]
            if len(noos) > 0 and noos.mean() > 0:
                avg_noos = noos.mean()
                adj = np.where(o == 1, avg_noos, s)
                for m in range(n):
                    if o[m] == 0 and p[m] > 0 and p[m] < avg_noos * 0.5:
                        adj[m] = 0.5 * s[m] + 0.5 * avg_noos
            else:
                adj = s.copy(); avg_noos = s.mean()

            if n >= 2:
                lev = adj[0]; tr = (adj[-1] - adj[0]) / max(n - 1, 1)
                for i in range(1, n):
                    nl = self.alpha * adj[i] + (1 - self.alpha) * (lev + tr)
                    nt = self.beta * (nl - lev) + (1 - self.beta) * tr
                    lev, tr = nl, nt
                holt = lev + tr
            else:
                holt = adj[0]; tr = 0

            w = WMA_WEIGHTS[-n:] if n <= 5 else WMA_WEIGHTS
            w = w / w.sum()
            wma = np.dot(adj[-len(w):], w) if n >= 3 else adj.mean()

            if n >= 4:
                comb = 0.5 * holt + 0.5 * wma
            else:
                comb = 0.5 * holt + 0.5 * wma

            mean_a = adj.mean()
            if mean_a > 0 and n >= 3:
                comb *= (1 + min((np.std(adj) / mean_a) * 0.3, 0.5))

            pred = max(0, comb)
            preds[(it['idk'], it['ida'])] = (pred, s.mean())

        # Smart rounding
        items = [{'k': k, 'p': v[0], 'a': v[1]} for k, v in preds.items()]
        df_p = pd.DataFrame(items)
        df_p['pr'] = df_p['p'].apply(lambda x: round(x) if x >= 0.5 else (1 if x > 0 else 0))

        for ida in df_p['k'].apply(lambda x: x[1]).unique():
            mask = df_p['k'].apply(lambda x: x[1] == ida)
            sub = df_p[mask]
            tgt = round(sub['p'].sum())
            cur = sub['pr'].sum()
            d = tgt - cur
            if d != 0:
                rem = sub['p'] - np.floor(sub['p'])
                adj_mask = sub['p'] > 0
                if d > 0:
                    for idx in rem[adj_mask].sort_values(ascending=False).index[:int(d)]:
                        df_p.loc[idx, 'pr'] += 1
                elif d < 0:
                    for idx in rem[adj_mask & (sub['pr'] > 0)].sort_values(ascending=True).index[:int(abs(d))]:
                        if df_p.loc[idx, 'pr'] > 0: df_p.loc[idx, 'pr'] -= 1

        df_p['ar'] = df_p['a'].apply(lambda x: round(x))
        self.pred_dict = {}
        for _, r in df_p.iterrows():
            k = r['k']
            self.pred_dict[k] = (int(r['pr']), int(r['ar']), int(r['pr'] - r['ar']))

        tp = sum(v[0] for v in self.pred_dict.values())
        self.log(f"🔮 Predikcija ukupno: {tp}")

    def _merge_lager(self):
        for _, k in self.all_keys.iterrows():
            idk, ida = k['ID KOMITENTA'], k['id artikla']
            pred, avg, razl = self.pred_dict.get((idk, ida), (0, 0, 0))
            lager = self.trenutni_dict.get((idk, ida), None)
            idx = self.df_monthly[
                (self.df_monthly['ID KOMITENTA']==idk) &
                (self.df_monthly['id artikla']==ida)].index
            if len(idx) > 0:
                ix = idx[0]
                self.df_monthly.loc[ix, 'Predikcija'] = pred
                self.df_monthly.loc[ix, 'Prosek'] = avg
                self.df_monthly.loc[ix, 'Razlika'] = razl
                if lager is not None:
                    self.df_monthly.loc[ix, 'Lager_danas'] = lager
                else:
                    lg, lm = self.meseci_order[-1]
                    _, lv, _ = self.prodaja_dict.get((idk, ida, lg, lm), (0, 0, 0))
                    self.df_monthly.loc[ix, 'Lager_danas'] = int(lv) if not pd.isna(lv) else 0

        for col in ['Predikcija','Prosek','Razlika','Lager_danas']:
            if col not in self.df_monthly.columns: self.df_monthly[col] = 0
            self.df_monthly[col] = self.df_monthly[col].fillna(0).astype(int)

    def _compute_orders(self):
        self.df_result = self.df_monthly.copy()

        def p1(row):
            if row['ID KOMITENTA'] in self.excluded: return 0
            return max(int(row['Predikcija']) - int(row['Lager_danas']), 0)

        def p2(row):
            if row['ID KOMITENTA'] in self.excluded: return 0
            pred, lag = int(row['Predikcija']), int(row['Lager_danas'])
            return max(max(pred - lag, 0), max(self.min_lager - lag, 0))

        self.df_result['Porudzbina_1'] = self.df_result.apply(p1, axis=1).astype(int)
        self.df_result['Porudzbina_2'] = self.df_result.apply(p2, axis=1).astype(int)

        t1 = self.df_result[~self.df_result['ID KOMITENTA'].isin(self.excluded)]['Porudzbina_1'].sum()
        t2 = self.df_result[~self.df_result['ID KOMITENTA'].isin(self.excluded)]['Porudzbina_2'].sum()
        self.log(f"📦 P1: {t1} | P2: {t2}")

    def _apply_min_order(self):
        self.adjustments = []
        for kid in sorted(self.df_result['ID KOMITENTA'].unique()):
            if kid in self.excluded: continue
            mask = self.df_result['ID KOMITENTA'] == kid
            total = self.df_result.loc[mask, 'Porudzbina_2'].sum()
            if 1 <= total < self.min_order:
                needed = self.min_order - total
                if total >= 2:
                    cands = self.df_result.loc[
                        mask & (self.df_result['Porudzbina_2'] > 0)
                    ].sort_values('Predikcija', ascending=False)
                    rem = int(needed)
                    for idx in cands.index:
                        if rem <= 0: break
                        add = min(rem, 2)
                        self.df_result.loc[idx, 'Porudzbina_2'] += add
                        rem -= add
                    if rem > 0 and len(cands) > 0:
                        self.df_result.loc[cands.index[0], 'Porudzbina_2'] += rem
                    self.adjustments.append((kid, f"{total}→{self.min_order}"))
                else:
                    self.df_result.loc[mask, 'Porudzbina_2'] = 0
                    self.adjustments.append((kid, f"{total}→0"))


# ──────────────────────────── EXCEL EXPORT ────────────────────────────

def create_excel(engine):
    df = engine.df_result
    wb = Workbook()

    hff = PatternFill('solid', fgColor='7C3AED')
    hfn = Font(bold=True, color='FFFFFF', name='Arial', size=10)
    tb = Border(left=Side('thin','E5E7EB'), right=Side('thin','E5E7EB'),
                top=Side('thin','E5E7EB'), bottom=Side('thin','E5E7EB'))
    ca = Alignment(horizontal='center', vertical='center')
    df_font = Font(name='Arial', size=9)

    ws = wb.active
    ws.title = "Porudžbina"
    headers = ['ID Komitenta','ID Artikla','Naziv artikla','Grupa',
               'Predikcija','Prosek','Lager danas','Porudžbina (osn.)','Porudžbina (min 2)']
    for c, h in enumerate(headers, 1):
        cell = ws.cell(1, c, h)
        cell.font = hfn; cell.fill = hff; cell.alignment = ca; cell.border = tb

    for i, row in df.iterrows():
        r = i + 2
        vals = [row['ID KOMITENTA'], row['id artikla'], row['Naziv artikla'],
                row['Grupa'], row['Predikcija'], row['Prosek'],
                row['Lager_danas'], row['Porudzbina_1'], row['Porudzbina_2']]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(r, c, v)
            cell.font = df_font; cell.border = tb
            if c in [1,2,5,6,7,8,9]: cell.alignment = ca
        if row['ID KOMITENTA'] in engine.excluded:
            for c in range(1, 10):
                ws.cell(r, c).fill = PatternFill('solid', fgColor='FFF2CC')
        elif row['Porudzbina_2'] > 0:
            ws.cell(r, 9).fill = PatternFill('solid', fgColor='E2EFDA')
            ws.cell(r, 9).font = Font(name='Arial', size=9, bold=True, color='2F5496')

    for col, w in zip('ABCDEFGHI', [14, 12, 52, 14, 12, 12, 12, 16, 16]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f'A1:I{len(df)+1}'

    # Sumarno
    ws2 = wb.create_sheet("Sumarno")
    for c, h in enumerate(['ID Komitenta','P1','P2','Artikala','Status'], 1):
        cell = ws2.cell(1, c, h)
        cell.font = hfn; cell.fill = hff; cell.alignment = ca; cell.border = tb
    sm = df.groupby('ID KOMITENTA').agg(
        p1=('Porudzbina_1','sum'), p2=('Porudzbina_2','sum'),
        na=('Porudzbina_2', lambda x: int((x>0).sum()))).reset_index()
    for ri, (_, sr) in enumerate(sm.sort_values('ID KOMITENTA').iterrows(), 2):
        kid = int(sr['ID KOMITENTA'])
        ws2.cell(ri,1,kid).font = df_font; ws2.cell(ri,1).alignment = ca
        ws2.cell(ri,2,int(sr['p1'])).font = df_font; ws2.cell(ri,2).alignment = ca
        ws2.cell(ri,3,int(sr['p2'])).font = df_font; ws2.cell(ri,3).alignment = ca
        ws2.cell(ri,4,int(sr['na'])).font = df_font; ws2.cell(ri,4).alignment = ca
        if kid in engine.excluded:
            ws2.cell(ri,5,'Isključen').font = df_font
            for c in range(1,6): ws2.cell(ri,c).fill = PatternFill('solid',fgColor='FFF2CC')
        elif sr['p2'] > 0:
            ws2.cell(ri,5,'Poručiti').font = Font(name='Arial',size=9,bold=True,color='2F5496')
        else:
            ws2.cell(ri,5,'-').font = df_font
        for c in range(1,6): ws2.cell(ri,c).border = tb
    for col, w in zip('ABCDE', [14,12,12,12,14]): ws2.column_dimensions[col].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ──────────────────────────── STREAMLIT APP ────────────────────────────

def main():
    # Header
    st.markdown("""
    <div class="header-banner">
        <div class="header-title">💨 VAPE PORUDŽBINE</div>
        <div class="header-sub">Predikcija prodaje & Automatsko generisanje porudžbina</div>
    </div>
    """, unsafe_allow_html=True)

    # Sidebar
    with st.sidebar:
        st.markdown("### ⚙️ Parametri")
        alpha = st.number_input("Alpha (nivo)", 0.0, 1.0, 0.4, 0.05,
                                help="Brzina reakcije na nove podatke")
        beta = st.number_input("Beta (trend)", 0.0, 1.0, 0.2, 0.05,
                               help="Brzina praćenja trenda")
        min_lager = st.number_input("Min lager", 0, 20, 2,
                                    help="Minimalan broj komada u radnji")
        min_order = st.number_input("Min porudžbina po objektu", 0, 50, 5,
                                    help="Ispod ovog broja se ne poručuje")

        st.markdown("---")
        st.markdown("### 🚫 Isključeni komitenti")
        excluded_str = st.text_area("Jedan ID po redu ili razdvojeni zarezom",
                                    value=DEFAULT_EXCLUDED, height=100)

    # Parse excluded
    excluded = set()
    for part in excluded_str.replace('\n', ',').split(','):
        part = part.strip()
        if part.isdigit():
            excluded.add(int(part))

    # Main content
    uploaded = st.file_uploader(
        "📂 Učitaj Excel fajl sa sheetovima: prodaja, startni lager, povrat, trenutni lager",
        type=['xlsx', 'xls'],
        help="Fajl mora imati sheet 'prodaja' i 'startni lager'. Sheetovi 'povrat' i 'trenutni lager' su opcioni."
    )

    if uploaded:
        file_bytes = uploaded.read()

        st.markdown(f"""
        <div class="success-box">
            ✅ <strong>{uploaded.name}</strong> učitan ({len(file_bytes)//1024} KB)
        </div>
        """, unsafe_allow_html=True)

        st.markdown("")

        if st.button("🚀  POKRENI — Generiši porudžbinu", use_container_width=True):
            progress_bar = st.progress(0)
            status = st.empty()

            try:
                engine = PredictionEngine(
                    file_bytes, excluded, alpha, beta, min_lager, min_order)
                result = engine.run(progress_bar, status)

                # Results
                st.markdown("---")

                total_pred = int(result['Predikcija'].sum())
                total_lager = int(result['Lager_danas'].sum())
                total_p1 = int(result[~result['ID KOMITENTA'].isin(excluded)]['Porudzbina_1'].sum())
                total_p2 = int(result[~result['ID KOMITENTA'].isin(excluded)]['Porudzbina_2'].sum())
                n_obj = result[result['Porudzbina_2'] > 0]['ID KOMITENTA'].nunique()

                # Metrics
                c1, c2, c3, c4, c5 = st.columns(5)
                with c1:
                    st.markdown(f'<div class="metric-card"><div class="metric-value">{total_pred}</div><div class="metric-label">Predikcija</div></div>', unsafe_allow_html=True)
                with c2:
                    st.markdown(f'<div class="metric-card"><div class="metric-value">{total_lager}</div><div class="metric-label">Trenutni lager</div></div>', unsafe_allow_html=True)
                with c3:
                    st.markdown(f'<div class="metric-card"><div class="metric-value">{total_p1}</div><div class="metric-label">P1 (osnovna)</div></div>', unsafe_allow_html=True)
                with c4:
                    st.markdown(f'<div class="metric-card"><div class="metric-value">{total_p2}</div><div class="metric-label">P2 (min {min_lager})</div></div>', unsafe_allow_html=True)
                with c5:
                    st.markdown(f'<div class="metric-card"><div class="metric-value">{n_obj}</div><div class="metric-label">Objekata</div></div>', unsafe_allow_html=True)

                st.markdown("")

                # Tabs
                tab1, tab2, tab3 = st.tabs(["📋 Porudžbina", "📊 Sumarno", "📝 Log"])

                with tab1:
                    display_cols = ['ID KOMITENTA', 'id artikla', 'Naziv artikla', 'Grupa',
                                    'Predikcija', 'Prosek', 'Lager_danas',
                                    'Porudzbina_1', 'Porudzbina_2']
                    show = result[display_cols].copy()
                    show.columns = ['ID Kom.', 'ID Art.', 'Naziv', 'Grupa',
                                    'Predikcija', 'Prosek', 'Lager', 'P1', 'P2']
                    st.dataframe(show, use_container_width=True, height=400)

                with tab2:
                    sm = result.groupby('ID KOMITENTA').agg(
                        P1=('Porudzbina_1', 'sum'),
                        P2=('Porudzbina_2', 'sum'),
                        Artikala=('Porudzbina_2', lambda x: int((x>0).sum()))
                    ).reset_index().sort_values('P2', ascending=False)
                    sm['Status'] = sm.apply(
                        lambda r: '🚫 Isklj.' if r['ID KOMITENTA'] in excluded
                        else ('✅ Poručiti' if r['P2'] > 0 else '—'), axis=1)
                    st.dataframe(sm, use_container_width=True, height=400)

                with tab3:
                    for msg in engine.logs:
                        st.text(msg)
                    if engine.adjustments:
                        st.markdown("**Korekcije min porudžbine:**")
                        for kid, note in engine.adjustments:
                            st.text(f"  Komitent {kid}: {note}")

                # Download
                st.markdown("---")
                excel_buf = create_excel(engine)
                fname = f"PORUDZBINA_{datetime.date.today().strftime('%Y%m%d')}.xlsx"

                st.download_button(
                    label=f"⬇️  Preuzmi Excel — {fname}",
                    data=excel_buf,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"❌ Greška: {str(e)}")
                import traceback
                st.code(traceback.format_exc())

    else:
        # Empty state
        st.markdown("""
        <div style="text-align: center; padding: 60px 20px; color: #aaa;">
            <div style="font-size: 48px; margin-bottom: 12px;">📂</div>
            <div style="font-size: 16px; color: #888;">Učitaj Excel fajl da počneš</div>
            <div style="font-size: 12px; color: #bbb; margin-top: 8px;">
                Potrebni sheetovi: prodaja, startni lager, povrat (opciono), trenutni lager (opciono)
            </div>
        </div>
        """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
