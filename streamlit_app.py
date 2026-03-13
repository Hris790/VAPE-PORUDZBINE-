import streamlit as st
import streamlit.components.v1 as components
import io, datetime, math, numpy as np, pandas as pd, json, requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter

# === GITHUB KONFIGURACIJA ===
GITHUB_RAW = "https://raw.githubusercontent.com/Hris790/VAPE-PORUDZBINE-/main/"
GITHUB_CONFIG = "https://raw.githubusercontent.com/Hris790/VAPE-PORUDZBINE-/main/config.json"

@st.cache_data(ttl=300)
def load_github_excel(filename):
    url = GITHUB_RAW + filename
    try:
        r = requests.get(url, timeout=15)
        if r.status_code == 200:
            return io.BytesIO(r.content)
        return None
    except:
        return None

@st.cache_data(ttl=60)
def load_github_config():
    try:
        r = requests.get(GITHUB_CONFIG, timeout=10)
        if r.status_code == 200:
            return r.json()
    except:
        pass
    return {"ukljuci_poslednji_mesec": False}

# --- PASSWORD ZASTITA ---
APP_PASSWORD = "vape2024"

def check_password():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if st.session_state.authenticated:
        return True

    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap');
    html, body, .stApp {
        background: #12002a !important;
        font-family: 'Poppins', sans-serif;
    }
    .stApp {
        background: linear-gradient(135deg, #12002a 0%, #1e0040 50%, #0d001f 100%) !important;
    }
    /* Sakrij streamlit elemente na login stranici */
    header[data-testid="stHeader"] { background: transparent !important; }
    .stDeployButton { display: none; }
    footer { display: none; }
    #MainMenu { display: none; }
    /* Centriraj login karticu */
    .block-container {
        max-width: 460px !important;
        margin: 0 auto !important;
        padding-top: 80px !important;
    }
    /* Input stilovi za tamnu pozadinu */
    .stTextInput > div > div > input {
        background: rgba(255,255,255,0.08) !important;
        border: 1px solid rgba(255,255,255,0.15) !important;
        color: white !important;
        border-radius: 12px !important;
        padding: 12px 16px !important;
        font-size: 15px !important;
    }
    .stTextInput > div > div > input::placeholder { color: rgba(255,255,255,0.35) !important; }
    .stTextInput > div > div > input:focus {
        border-color: rgba(168,85,247,0.6) !important;
        box-shadow: 0 0 0 3px rgba(168,85,247,0.15) !important;
    }
    /* Dugme */
    .stButton > button {
        background: linear-gradient(135deg, #a855f7 0%, #ec4899 100%) !important;
        color: white !important;
        border: none !important;
        border-radius: 12px !important;
        padding: 13px 32px !important;
        font-weight: 700 !important;
        font-size: 15px !important;
        width: 100% !important;
        box-shadow: 0 4px 20px rgba(168,85,247,0.35) !important;
        transition: opacity 0.2s !important;
    }
    .stButton > button:hover { opacity: 0.88 !important; }
    /* Error poruka */
    .stAlert { border-radius: 10px !important; background: rgba(220,38,38,0.15) !important; border: 1px solid rgba(220,38,38,0.3) !important; color: #fca5a5 !important; }
    </style>
    """, unsafe_allow_html=True)

    # Logo i naslov
    st.markdown("""
    <div style="text-align:center; margin-bottom: 36px;">
        <div style="display:inline-flex; align-items:center; gap:10px; margin-bottom: 24px;">
            <div style="width:36px; height:36px; background:linear-gradient(135deg,#a855f7,#ec4899);
                border-radius:9px; display:inline-flex; align-items:center; justify-content:center;">
                <div style="width:12px; height:12px; background:white; border-radius:3px; opacity:0.95;"></div>
            </div>
            <span style="font-size:22px; font-weight:700; color:white; letter-spacing:0.5px;">VAPE</span>
            <span style="font-size:22px; font-weight:300; color:rgba(255,255,255,0.45);">Analitika</span>
        </div>
        <div style="height:1px; background:linear-gradient(90deg, transparent, rgba(255,255,255,0.12), transparent); margin-bottom:28px;"></div>
        <h2 style="color:white; font-size:24px; font-weight:700; margin:0 0 8px 0; line-height:1.35;">
            Dobrodošli u aplikaciju<br>Vape Shop-a!
        </h2>
        <p style="color:rgba(255,255,255,0.4); font-size:14px; margin:0;">
            Unesite šifru za pristup sistemu
        </p>
    </div>
    """, unsafe_allow_html=True)

    pwd = st.text_input("Šifra", type="password", placeholder="Unesite šifru...", label_visibility="collapsed")
    btn = st.button("Prijavi se", use_container_width=True)
    if btn:
        if pwd == APP_PASSWORD:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Pogrešna šifra")

    st.markdown("""
    <div style="text-align:center; margin-top:28px;">
        <p style="color:rgba(255,255,255,0.18); font-size:12px; margin:0;">
            AMAN d.o.o. · Analitički sistem
        </p>
    </div>
    """, unsafe_allow_html=True)

    return False

if not check_password():
    st.stop()
# --- KRAJ PASSWORD ZASTITE ---

WMA_WEIGHTS = np.array([0.03, 0.07, 0.12, 0.28, 0.50])
HIST_WEIGHT = 0.03

class PredictionEngine:
    def __init__(self, file_bytes, excluded_ids, alpha, beta, min_lager, min_order, mesecni_trosak=0, analitika_meseci=None):
        self.file_bytes = file_bytes; self.excluded = excluded_ids
        self.alpha = alpha; self.beta = beta; self.min_lager = min_lager; self.min_order = min_order
        self.mesecni_trosak = mesecni_trosak
        self.analitika_meseci = analitika_meseci
        self.logs = []; self.adjustments = []; self.has_history = False
        self.has_prices = False

    def log(self, msg): self.logs.append(msg)

    def run(self, progress_bar):
        progress_bar.progress(5, "Ucitavanje..."); self._load_sheets()
        progress_bar.progress(15, "Priprema..."); self._prepare_lookups()
        progress_bar.progress(25, "Povrat/korekcija..."); self._compute_povrat()
        progress_bar.progress(40, "Mesecni pregled..."); self._build_monthly()
        progress_bar.progress(55, "Predikcija..."); self._predict_all()
        progress_bar.progress(70, "Lager..."); self._merge_lager()
        progress_bar.progress(80, "Porudzbine..."); self._compute_orders()
        progress_bar.progress(85, "Analitika..."); self._apply_min_order(); self._compute_analytics()
        progress_bar.progress(100, "Gotovo!"); return self.df_result

    def _load_sheets(self):
        xls = pd.ExcelFile(io.BytesIO(self.file_bytes))
        sm = {s.strip().lower(): s for s in xls.sheet_names}
        def find(kws):
            for kw in kws:
                for nl, no in sm.items():
                    if kw in nl: return no
            return None
        s_prod=find(['prodaja']); s_start=find(['startni']); s_pov=find(['povrat'])
        s_tl=find(['trenutni']); s_hist=find(['pre sept','pre sep','istorij'])
        if not s_prod: raise ValueError("Nema sheeta 'prodaja'!")
        if not s_start: raise ValueError("Nema sheeta 'startni lager'!")
        self.prodaja = pd.read_excel(xls, sheet_name=s_prod); self.prodaja.columns=[c.strip() for c in self.prodaja.columns]
        self.prodaja = self.prodaja[[c for c in self.prodaja.columns if 'Unnamed' not in str(c)]]
        self.log(f"Prodaja: {len(self.prodaja)} redova")
        self.region_map = {}
        if 'Region' in self.prodaja.columns:
            self.region_map = self.prodaja.drop_duplicates('ID KOMITENTA').set_index('ID KOMITENTA')['Region'].to_dict()
            self.log(f"Region: {len(set(self.region_map.values()))} regiona")
        self.startni = pd.read_excel(xls, sheet_name=s_start); self.startni.columns=[c.strip() for c in self.startni.columns]
        self.log(f"Startni: {len(self.startni)} redova")
        price_cols = ['Redovna cena','Akcijska cena','Finalna cena','Nabavna vrednost','Profit']
        if all(c in self.prodaja.columns for c in price_cols):
            self.has_prices = True; self.log("Cene i profit: DA")
        self.povrat_df = pd.DataFrame()
        if s_pov:
            self.povrat_df = pd.read_excel(xls, sheet_name=s_pov); self.povrat_df.columns=[c.strip() for c in self.povrat_df.columns]
            self.log(f"Povrat: {len(self.povrat_df)} redova")
        self.trenutni = pd.DataFrame()
        if s_tl:
            self.trenutni = pd.read_excel(xls, sheet_name=s_tl); self.trenutni.columns=[c.strip() for c in self.trenutni.columns]
            self.log(f"Trenutni lager: {len(self.trenutni)} redova")
        self.hist_df = pd.DataFrame()
        self.has_history = False
        _meseci_u_prodaji = self.prodaja[['Godina','Mesec']].drop_duplicates().values.tolist()
        _ima_pre_sept = any((int(g) < 2025) or (int(g) == 2025 and int(m) < 9) for g, m in _meseci_u_prodaji)
        if _ima_pre_sept:
            self.log("Rezim: KOMPLETAN ISTORIJAT u prodaja sheetu — istorijski sheet se ignorise")
        elif s_hist:
            self.hist_df = pd.read_excel(xls, sheet_name=s_hist); self.hist_df.columns=[c.strip() for c in self.hist_df.columns]
            self.has_history = True; self.log(f"Istorija: {len(self.hist_df)} redova")
        self.meseci_order = sorted(self.prodaja[['Godina','Mesec']].drop_duplicates().values.tolist())
        mn={1:'Jan',2:'Feb',3:'Mar',4:'Apr',5:'Maj',6:'Jun',7:'Jul',8:'Avg',9:'Sep',10:'Okt',11:'Nov',12:'Dec'}
        self.mesec_labels = [f"{mn.get(int(m),'?')} {int(g)}" for g,m in self.meseci_order]
        lg,lm = self.meseci_order[-1]; nm=int(lm)+1; ng=int(lg)
        if nm>12: nm=1; ng+=1
        self.pred_label = f"{mn.get(nm,'?')} {ng}"
        om=nm+1; og=ng
        if om>12: om=1; og+=1
        self.order_label = f"{mn.get(om,'?')} {og}"
        self.log(f"Meseci: {', '.join(self.mesec_labels)}")
        self.num_komitenti = self.prodaja['ID KOMITENTA'].nunique()
        self.trosak_po_objektu = self.mesecni_trosak / max(self.num_komitenti, 1) if self.mesecni_trosak > 0 else 0
        if self.mesecni_trosak > 0:
            self.log(f"Ukupan trosak: {self.mesecni_trosak:,.0f} / {self.num_komitenti} objekata = {self.trosak_po_objektu:,.0f} po objektu za period")

    def _prepare_lookups(self):
        kp = self.prodaja[['ID KOMITENTA','id artikla','Naziv artikla','Grupa']].drop_duplicates()
        ks = self.startni[['ID KOMITENTA','id artikla','Naziv artikla','Grupa']].drop_duplicates()
        frames = [kp, ks]
        if self.has_history:
            hcols = self.hist_df.columns.tolist()
            col_map = {}
            for c in hcols:
                cl = c.lower()
                if 'komitent' in cl: col_map[c] = 'ID KOMITENTA'
                elif 'id' in cl and 'artikl' in cl: col_map[c] = 'id artikla'
                elif 'naziv' in cl and 'artikl' in cl: col_map[c] = 'Naziv artikla'
                elif 'grup' in cl: col_map[c] = 'Grupa'
            hdf = self.hist_df.rename(columns=col_map)
            for nc in ['ID KOMITENTA','id artikla','Naziv artikla','Grupa']:
                if nc not in hdf.columns: hdf[nc] = ''
            kh = hdf[['ID KOMITENTA','id artikla','Naziv artikla','Grupa']].drop_duplicates()
            frames.append(kh)
        self.all_keys = pd.concat(frames).drop_duplicates().sort_values(['ID KOMITENTA','id artikla']).reset_index(drop=True)
        self.startni_dict = {(r['ID KOMITENTA'],r['id artikla']): r['Kolicina'] for _,r in self.startni.iterrows()}
        self.has_promet = 'PROMET KA NJIMA' in self.prodaja.columns
        self.prodaja_dict = {}
        for _,r in self.prodaja.iterrows():
            key=(r['ID KOMITENTA'],r['id artikla'],r['Godina'],r['Mesec'])
            pm = r['PROMET KA NJIMA'] if self.has_promet else 0
            self.prodaja_dict[key] = (r.get('Prodata Kolicina',r.get('Kolicina',0)), r.get('Lager',0), pm if not pd.isna(pm) else 0)
        self.hist_dict={}; self.hist_total_dict={}; self.hist_months_per_art={}
        if self.has_history:
            ha = self.hist_df.groupby(['ID KOMITENTA','id artikla'])['Prodata Kolicina'].agg(['sum','mean']).reset_index()
            for _,r in ha.iterrows():
                self.hist_dict[(int(r['ID KOMITENTA']),int(r['id artikla']))] = float(r['mean'])
                self.hist_total_dict[(int(r['ID KOMITENTA']),int(r['id artikla']))] = int(r['sum'])
            for ida in self.hist_df['id artikla'].unique():
                sub=self.hist_df[self.hist_df['id artikla']==ida]
                self.hist_months_per_art[int(ida)]=sub[['Godina','Mesec']].drop_duplicates().shape[0]
            self.log(f"Istorijski prosek za {len(self.hist_dict)} kombinacija")
        self.recent_months_per_art={}
        for ida in self.prodaja['id artikla'].unique():
            sub=self.prodaja[self.prodaja['id artikla']==ida]
            self.recent_months_per_art[int(ida)]=sub[['Godina','Mesec']].drop_duplicates().shape[0]
        self.total_months_per_art={}
        all_arts=set([int(x) for x in self.prodaja['id artikla'].unique()])
        if self.has_history: all_arts|=set([int(x) for x in self.hist_df['id artikla'].unique()])
        for ida in all_arts:
            self.total_months_per_art[ida]=self.hist_months_per_art.get(ida,0)+self.recent_months_per_art.get(ida,0)
        self.povrat_total={}
        if len(self.povrat_df)>0:
            ic=[c for c in self.povrat_df.columns if 'id' in c.lower() and 'artikl' in c.lower()]
            mc=[c for c in self.povrat_df.columns if 'mesec' in c.lower()]
            gc=[c for c in self.povrat_df.columns if 'godin' in c.lower()]
            kc=[c for c in self.povrat_df.columns if 'koli' in c.lower()]
            if ic and mc and gc and kc:
                for _,r in self.povrat_df.iterrows():
                    key=(r[ic[0]],r[gc[0]],r[mc[0]]); self.povrat_total[key]=self.povrat_total.get(key,0)+r[kc[0]]
        self.trenutni_dict={}
        if len(self.trenutni)>0:
            ikc=[c for c in self.trenutni.columns if 'komitent' in c.lower()]
            iac=[c for c in self.trenutni.columns if 'artikl' in c.lower() and 'id' in c.lower()]
            lc=[c for c in self.trenutni.columns if 'lager' in c.lower()]
            if ikc and iac and lc:
                for _,r in self.trenutni.iterrows():
                    k,a=r[ikc[0]],r[iac[0]]
                    if pd.notna(k) and pd.notna(a): self.trenutni_dict[(int(k),int(a))]=int(r[lc[0]]) if pd.notna(r[lc[0]]) else 0
        self.profit_per_unit = {}
        self.price_info = {}
        if self.has_prices:
            for ida in self.prodaja['id artikla'].unique():
                sub = self.prodaja[self.prodaja['id artikla']==ida].iloc[0]
                red, akc, fin, nab = sub['Redovna cena'], sub['Akcijska cena'], sub['Finalna cena'], sub['Nabavna vrednost']
                ppu_fin = fin/1.2/1.2 - nab
                ppu_red = red/1.2/1.2 - nab
                self.profit_per_unit[int(ida)] = ppu_fin
                self.price_info[int(ida)] = {'redovna': red, 'akcijska': akc, 'finalna': fin, 'nabavna': nab, 'profit_akcija': ppu_fin, 'profit_redovna': ppu_red}
        self.log(f"Kombinacija: {len(self.all_keys)}")

    def _compute_povrat(self):
        self.final_povrat={}; self.final_korekcija={}
        if not self.has_promet or not self.povrat_total: return
        implied={}
        for _,k in self.all_keys.iterrows():
            idk,ida=k['ID KOMITENTA'],k['id artikla']; poc=self.startni_dict.get((idk,ida),0)
            for god,mes in self.meseci_order:
                pv,lv,tv=self.prodaja_dict.get((idk,ida,god,mes),(0,0,0)); lv=lv if not pd.isna(lv) else 0
                implied[(idk,ida,god,mes)]=poc+tv-pv-lv; poc=lv
        all_art=set(list(self.prodaja['id artikla'].unique())+list(self.startni['id artikla'].unique()))
        for god,mes in self.meseci_order:
            for ida in all_art:
                ap=self.povrat_total.get((ida,god,mes),0); pi={}; ni={}
                for _,k in self.all_keys[self.all_keys['id artikla']==ida].iterrows():
                    i2=k['ID KOMITENTA']; im=implied.get((i2,ida,god,mes),0)
                    if im>0: pi[i2]=im
                    elif im<0: ni[i2]=im
                tp=sum(pi.values())
                if ap>0 and tp>0:
                    raw={i:ap*(v/tp) for i,v in pi.items()}; fl={i:int(v) for i,v in raw.items()}
                    d=ap-sum(fl.values()); rem={i:raw[i]-fl[i] for i in raw}
                    for j,i in enumerate(sorted(rem,key=rem.get,reverse=True)):
                        if j<int(d): fl[i]+=1
                    for i,pv2 in fl.items():
                        self.final_povrat[(i,ida,god,mes)]=pv2; self.final_korekcija[(i,ida,god,mes)]=pi[i]-pv2
                elif ap==0:
                    for i,v in pi.items(): self.final_korekcija[(i,ida,god,mes)]=v
                for i,v in ni.items(): self.final_korekcija[(i,ida,god,mes)]=self.final_korekcija.get((i,ida,god,mes),0)+v

    def _build_monthly(self):
        rows=[]
        for _,k in self.all_keys.iterrows():
            idk,ida=k['ID KOMITENTA'],k['id artikla']; poc=self.startni_dict.get((idk,ida),0)
            row={'ID KOMITENTA':idk,'id artikla':ida,'Naziv artikla':k['Naziv artikla'],'Grupa':k['Grupa']}
            row['Total_JanAvg']=self.hist_total_dict.get((idk,ida),0)
            for i,(god,mes) in enumerate(self.meseci_order):
                lb=self.mesec_labels[i]; pv,lv,tv=self.prodaja_dict.get((idk,ida,god,mes),(0,0,0))
                lv=lv if not pd.isna(lv) else 0; tv=tv if not pd.isna(tv) else 0
                row[f'{lb}_Pocetno']=poc; row[f'{lb}_Promet']=tv; row[f'{lb}_Prodaja']=pv
                row[f'{lb}_Povrat']=self.final_povrat.get((idk,ida,god,mes),0)
                row[f'{lb}_Korekcija']=self.final_korekcija.get((idk,ida,god,mes),0); poc=lv
            rows.append(row)
        self.df_monthly=pd.DataFrame(rows)

    def _predict_all(self):
        analysis=[]
        for _,k in self.all_keys.iterrows():
            idk,ida=k['ID KOMITENTA'],k['id artikla']; poc=self.startni_dict.get((idk,ida),0)
            sales,oos,pocs,end_lagers,promets=[],[],[],[],[]
            for god,mes in self.meseci_order:
                pv,lv,tv=self.prodaja_dict.get((idk,ida,god,mes),(0,0,0))
                lv=lv if not pd.isna(lv) else 0; tv=tv if not pd.isna(tv) else 0
                sales.append(pv); oos.append(1 if poc==0 else 0); pocs.append(poc)
                end_lagers.append(lv); promets.append(tv); poc=lv
            ha=self.hist_dict.get((idk,ida),0)
            lager_danas=self.trenutni_dict.get((idk,ida),0)
            analysis.append({'idk':idk,'ida':ida,'sales':np.array(sales,dtype=float),'oos':np.array(oos),
                'poc':np.array(pocs,dtype=float),'ha':ha,'lager_danas':lager_danas,
                'end_lagers':np.array(end_lagers,dtype=float),'promets':np.array(promets,dtype=float)})
        preds={}
        for it in analysis:
            s,o,p=it['sales'],it['oos'],it['poc']; n=len(s); ha=it['ha']
            lager_danas=it['lager_danas']
            el=it['end_lagers']; tv=it['promets']
            constrained = np.zeros(n, dtype=bool)
            for m in range(n):
                if p[m]==0 and tv[m]==0: constrained[m] = True
                elif el[m]==0 and s[m]>0: constrained[m] = True
                elif p[m]==0 and tv[m]>0 and el[m]==0: constrained[m] = True
            normal_mask = ~constrained & (p > 0)
            normal_sales = s[normal_mask]
            normal_with_sales = normal_sales[normal_sales > 0]
            if len(normal_with_sales) > 0: an = normal_with_sales.mean()
            elif len(normal_sales) > 0: an = normal_sales.mean()
            else: an = 0
            if an > 0:
                adj = s.copy().astype(float)
                for m in range(n):
                    if constrained[m]:
                        if p[m]==0 and tv[m]==0: adj[m] = an
                        elif el[m]==0 and s[m]>0: adj[m] = max(an, s[m])
                        else: adj[m] = an
                    elif p[m]>0 and p[m]<an*0.5: adj[m] = 0.5*s[m] + 0.5*an
            elif ha>0: adj=np.full(n,ha)
            else: adj=s.copy().astype(float)
            if n>=2:
                lev=adj[0]; tr=(adj[-1]-adj[0])/max(n-1,1)
                for i in range(1,n):
                    nl=self.alpha*adj[i]+(1-self.alpha)*(lev+tr); nt=self.beta*(nl-lev)+(1-self.beta)*tr; lev,tr=nl,nt
                holt=lev+tr
            else: holt=adj[0]
            w=WMA_WEIGHTS[-n:] if n<=5 else WMA_WEIGHTS; w=w/w.sum()
            wma=np.dot(adj[-len(w):],w) if n>=3 else adj.mean()
            comb = 0.4 * min(holt, wma) + 0.6 * max(holt, wma)
            ma=adj.mean()
            if ma>0 and n>=3: comb*=(1+min((np.std(adj)/ma)*0.4,0.7))
            if ha>0 and comb>0: comb=(1-HIST_WEIGHT)*comb+HIST_WEIGHT*ha
            elif ha>0 and comb==0 and s.sum()==0: comb=ha*0.20
            has_recent_sales = (s[-2:].sum() > 0) if n >= 2 else (s.sum() > 0)
            if lager_danas <= 2 and has_recent_sales:
                stocked_sales = [s[i] for i in range(n) if p[i] > 0]
                avg_when_stocked = np.mean(stocked_sales) if stocked_sales else 0
                if avg_when_stocked > 0 and comb < avg_when_stocked: comb = avg_when_stocked
            if ma > 5 and comb < ma: comb = ma
            avg_5m_raw = float(adj[-5:].mean()) if n >= 5 else float(adj.mean())
            ht=self.hist_total_dict.get((it['idk'],it['ida']),0)
            rt=float(s.sum()); tm=self.total_months_per_art.get(it['ida'],n)
            full_avg=(ht+rt)/max(tm,1)
            if comb < full_avg and comb > 0:
                if n >= 5: declining = all(adj[i] <= adj[i-1] for i in range(n-4, n))
                elif n >= 3: declining = all(adj[i] <= adj[i-1] for i in range(1, n))
                else: declining = (n >= 2 and adj[-1] <= adj[-2])
                if not declining: comb = full_avg
            if comb <= 0:
                last5 = s[-5:] if n >= 5 else s
                if last5.sum() > 0:
                    comb = 1.0
                    if s[-1] > 1: comb = s[-1]
            preds[(it['idk'],it['ida'])]=(max(0,comb),full_avg,avg_5m_raw)
        items=[{'k':k,'p':v[0],'a':v[1],'avg5':v[2]} for k,v in preds.items()]; df_p=pd.DataFrame(items)
        df_p['pr']=df_p['p'].apply(lambda x: round(x))
        df_p['ar']=df_p['a'].apply(lambda x: round(x))
        self.pred_dict={r['k']:(int(r['pr']),int(r['ar']),int(r['pr']-r['ar']),r['avg5']) for _,r in df_p.iterrows()}
        self.log(f"Predikcija: {sum(v[0] for v in self.pred_dict.values())} kom")

    def _merge_lager(self):
        for _,k in self.all_keys.iterrows():
            idk,ida=k['ID KOMITENTA'],k['id artikla']; pred,avg,razl,avg5m=self.pred_dict.get((idk,ida),(0,0,0,0))
            lager=self.trenutni_dict.get((idk,ida),None)
            idx=self.df_monthly[(self.df_monthly['ID KOMITENTA']==idk)&(self.df_monthly['id artikla']==ida)].index
            if len(idx)>0:
                ix=idx[0]; self.df_monthly.loc[ix,'Predikcija']=pred; self.df_monthly.loc[ix,'Prosek']=avg; self.df_monthly.loc[ix,'Razlika']=razl
                self.df_monthly.loc[ix,'Avg5m']=avg5m
                if lager is not None: self.df_monthly.loc[ix,'Lager_danas']=lager
                else: self.df_monthly.loc[ix,'Lager_danas']=0
        for col in ['Predikcija','Prosek','Razlika','Lager_danas']:
            if col not in self.df_monthly.columns: self.df_monthly[col]=0
            self.df_monthly[col]=self.df_monthly[col].fillna(0).astype(int)
        if 'Avg5m' not in self.df_monthly.columns: self.df_monthly['Avg5m']=0
        self.df_monthly['Avg5m']=self.df_monthly['Avg5m'].fillna(0)

    def _compute_orders(self):
        self.df_result=self.df_monthly.copy()
        def p1(row):
            if row['ID KOMITENTA'] in self.excluded: return 0
            return max(int(row['Predikcija'])-int(row['Lager_danas']),0)
        def p2(row):
            if row['ID KOMITENTA'] in self.excluded: return 0
            pred=int(row['Predikcija']); lager=int(row['Lager_danas']); prosek=int(row['Prosek'])
            osnova=max(pred-lager,0)
            if self.min_lager is not None and lager < self.min_lager and pred > 0:
                dopuna = max(self.min_lager - lager, osnova)
            else:
                dopuna = osnova
            return dopuna
        self.df_result['Porudzbina_1']=self.df_result.apply(p1,axis=1).astype(int)
        self.df_result['Porudzbina_2']=self.df_result.apply(p2,axis=1).astype(int)

        last_label = self.mesec_labels[-1]

        def extra_buffer(prodaja_poslednji):
            if prodaja_poslednji <= 0: return 0
            elif prodaja_poslednji <= 5: return 2
            elif prodaja_poslednji <= 10: return 3
            elif prodaja_poslednji <= 15: return 4
            else: return 5

        def finalna_provera(row):
            if row['ID KOMITENTA'] in self.excluded: return int(row['Porudzbina_2'])
            p2_val = int(row['Porudzbina_2'])
            lager = int(row['Lager_danas'])
            prodaja_poslednji = int(row.get(f'{last_label}_Prodaja', 0))
            if (p2_val + lager) <= prodaja_poslednji:
                dodatak = extra_buffer(prodaja_poslednji)
                return p2_val + dodatak
            return p2_val

        self.df_result['Porudzbina_2'] = self.df_result.apply(finalna_provera, axis=1).astype(int)
        n_korigovano = (self.df_result['Porudzbina_2'] > self.df_result.apply(p2, axis=1)).sum()
        self.log(f"Finalna provera P2: {n_korigovano} kombinacija korigovano (porudzbina+lager <= prodaja poslednjeg meseca)")

    def _apply_min_order(self):
        self.adjustments = []
        if self.min_order is None or self.min_order <= 0: return

        grp = self.df_result.groupby('ID KOMITENTA')['Porudzbina_2'].sum()
        ima_nesto = grp[grp > 0]
        granica = self.min_order / 2

        premali = ima_nesto[ima_nesto < granica].index
        dopuni = ima_nesto[(ima_nesto >= granica) & (ima_nesto < self.min_order)].index

        mask_gasi = self.df_result['ID KOMITENTA'].isin(premali)
        n_gasi = len(premali)
        self.df_result.loc[mask_gasi, 'Porudzbina_2'] = 0

        n_dopuni = 0
        for komt_id in dopuni:
            mask_obj = (self.df_result['ID KOMITENTA'] == komt_id) & (self.df_result['Porudzbina_2'] > 0)
            ukupno = int(self.df_result.loc[self.df_result['ID KOMITENTA'] == komt_id, 'Porudzbina_2'].sum())
            nedostaje = self.min_order - ukupno
            if nedostaje <= 0 or not mask_obj.any(): continue
            idx_max = self.df_result.loc[mask_obj, 'Porudzbina_2'].idxmax()
            self.df_result.at[idx_max, 'Porudzbina_2'] += nedostaje
            n_dopuni += 1

        if n_gasi > 0:
            self.log(f"Min order ({self.min_order} kom): {n_gasi} objekata imalo premalo komada ukupno — postavljeno na 0")
        if n_dopuni > 0:
            self.log(f"Min order ({self.min_order} kom): {n_dopuni} objekata dopunjeno do minimuma {self.min_order} kom")

    def _compute_analytics(self):
        if not self.has_prices:
            self.df_oos = pd.DataFrame()
            self.df_profit_obj = pd.DataFrame()
            self.df_promo = pd.DataFrame()
            self.analitika_labels = []
            return
        df = self.df_result; ml = self.mesec_labels

        if self.analitika_meseci and len(self.analitika_meseci) > 0:
            a_meseci = self.analitika_meseci
        else:
            a_meseci = self.meseci_order

        a_indices = []
        for i, (g, m) in enumerate(self.meseci_order):
            for ag, am in a_meseci:
                if int(g) == int(ag) and int(m) == int(am):
                    a_indices.append(i); break

        if not a_indices:
            a_indices = list(range(len(self.meseci_order)))

        a_labels = [ml[i] for i in a_indices]
        a_meseci_order = [self.meseci_order[i] for i in a_indices]
        n_a = len(a_indices)
        self.analitika_labels = a_labels
        self.log(f"Analitika period: {', '.join(a_labels)} ({n_a} meseci)")

        a_set = set((int(g), int(m)) for g, m in a_meseci_order)
        prodaja_a = self.prodaja[self.prodaja.apply(lambda r: (int(r['Godina']), int(r['Mesec'])) in a_set, axis=1)]

        ppu_mesec = {}
        if self.has_prices:
            for (ida_v, god_v, mes_v), grp in self.prodaja.groupby(['id artikla','Godina','Mesec']):
                kol = grp['Prodata Kolicina'].sum()
                if kol > 0:
                    ppu_mesec[(int(ida_v), int(god_v), int(mes_v))] = grp['Profit'].sum() / kol
                else:
                    r0 = grp.iloc[0]
                    ppu_mesec[(int(ida_v), int(god_v), int(mes_v))] = r0['Finalna cena'] / 1.2 / 1.2 - r0['Nabavna vrednost']

        def get_ppu(ida_v, god_v, mes_v):
            key = (int(ida_v), int(god_v), int(mes_v))
            if key in ppu_mesec:
                return ppu_mesec[key]
            art_keys = sorted([k for k in ppu_mesec if k[0] == int(ida_v)], key=lambda x: (x[1], x[2]))
            if not art_keys:
                return self.profit_per_unit.get(int(ida_v), 0)
            target = int(god_v) * 12 + int(mes_v)
            best = min(art_keys, key=lambda x: abs(x[1] * 12 + x[2] - target))
            return ppu_mesec[best]

        oos_rows = []
        for _, k in self.all_keys.iterrows():
            idk, ida = k['ID KOMITENTA'], k['id artikla']
            poc = self.startni_dict.get((idk, ida), 0)
            month_sales = []; month_oos = []; month_gm = []
            for i, (god, mes) in enumerate(self.meseci_order):
                lb = ml[i]
                pv = df[(df['ID KOMITENTA']==idk)&(df['id artikla']==ida)][f'{lb}_Prodaja'].values
                pv = int(pv[0]) if len(pv) > 0 else 0
                if i in a_indices:
                    month_sales.append(pv)
                    month_oos.append(poc == 0)
                    month_gm.append((god, mes))
                lv_col = self.prodaja_dict.get((idk, ida, god, mes), (0, 0, 0))
                poc = lv_col[1] if not pd.isna(lv_col[1]) else 0
            non_oos_sales = [month_sales[j] for j in range(len(month_sales)) if not month_oos[j]]
            avg_stocked = np.mean(non_oos_sales) if non_oos_sales else 0
            oos_count = sum(month_oos)
            if oos_count > 0 and avg_stocked > 0:
                row = {
                    'ID KOMITENTA': idk, 'id artikla': ida,
                    'Naziv artikla': k['Naziv artikla'], 'Grupa': k['Grupa'],
                    'Prosek_kad_ima': round(avg_stocked, 1),
                    'Lager_danas': self.trenutni_dict.get((idk, ida), 0)
                }
                total_lost = 0
                for j in range(len(month_sales)):
                    god_j, mes_j = month_gm[j]
                    lb_j = a_labels[j]
                    if month_oos[j]:
                        ppu_j = get_ppu(ida, god_j, mes_j)
                        izgub_j = round(avg_stocked * ppu_j, 0)
                        row[f'OOS_{lb_j}'] = 1
                        row[f'Izgub_{lb_j}'] = izgub_j
                        total_lost += izgub_j
                    else:
                        row[f'OOS_{lb_j}'] = 0
                        row[f'Izgub_{lb_j}'] = 0
                row['OOS_meseci'] = oos_count
                row['Izgubljeni_profit'] = round(total_lost, 0)
                oos_rows.append(row)
        self.df_oos = pd.DataFrame(oos_rows)
        if len(self.df_oos) > 0:
            self.df_oos = self.df_oos.sort_values('Izgubljeni_profit', ascending=False)
            self.log(f"OOS analiza: {len(self.df_oos)} kombinacija, izgubljeno {self.df_oos['Izgubljeni_profit'].sum():,.0f} RSD")

        trosak_mes_po_obj = self.trosak_po_objektu / max(n_a, 1) if self.trosak_po_objektu > 0 else 0

        profit_rows = []
        for idk in self.prodaja['ID KOMITENTA'].unique():
            sub = prodaja_a[prodaja_a['ID KOMITENTA'] == idk]
            total_prod = int(sub['Prodata Kolicina'].sum())
            total_profit = sub['Profit'].sum()
            n_art = self.all_keys[self.all_keys['ID KOMITENTA'] == idk]['id artikla'].nunique()
            mes_data = {}
            for _, r in sub.iterrows():
                key = f"{int(r['Godina'])}/{int(r['Mesec'])}"
                mes_data[key] = mes_data.get(key, 0) + r['Profit']
            mes_data_neto = {k: v - trosak_mes_po_obj for k, v in mes_data.items()}
            oos_sub = self.df_oos[self.df_oos['ID KOMITENTA'] == idk] if len(self.df_oos) > 0 else pd.DataFrame()
            lost = oos_sub['Izgubljeni_profit'].sum() if len(oos_sub) > 0 else 0
            trosak_total = self.trosak_po_objektu
            neto = total_profit - trosak_total
            row_dict = {
                'ID KOMITENTA': int(idk), 'Artikala': n_art,
                'Prodato_kom': total_prod, 'Bruto_profit': round(total_profit, 0),
                'Trosak_mkt': round(trosak_total, 0),
                'Neto_profit': round(neto, 0),
                'Izgubljeno_OOS': round(lost, 0),
                'Potencijalni_profit': round(neto + lost, 0),
            }
            for j in range(n_a):
                key_j = f"{int(a_meseci_order[j][0])}/{int(a_meseci_order[j][1])}"
                row_dict[f'Neto_{a_labels[j]}'] = round(mes_data_neto.get(key_j, -trosak_mes_po_obj), 0)
                row_dict[f'Bruto_{a_labels[j]}'] = round(mes_data.get(key_j, 0), 0)
            profit_rows.append(row_dict)
        self.trosak_mes_po_obj = trosak_mes_po_obj
        self.df_profit_obj = pd.DataFrame(profit_rows).sort_values('Neto_profit', ascending=True)

        promo_rows = []
        for ida in self.prodaja['id artikla'].unique():
            pi = self.price_info.get(int(ida), {})
            if not pi: continue
            sub = prodaja_a[prodaja_a['id artikla'] == ida]
            total_prod = int(sub['Prodata Kolicina'].sum())
            if total_prod == 0: continue
            profit_akcija = sub['Profit'].sum()
            profit_redovna = pi['profit_redovna'] * total_prod
            razlika = profit_redovna - profit_akcija
            prihod_akcija = (sub['Finalna cena'] * sub['Prodata Kolicina']).sum()
            prihod_redovna = (sub['Redovna cena'] * sub['Prodata Kolicina']).sum()
            first_a_idx = a_indices[0]
            if first_a_idx == 0:
                start_lager = self.startni[self.startni['id artikla']==ida]['Kolicina'].sum() if 'Kolicina' in self.startni.columns else 0
            else:
                prev_god, prev_mes = self.meseci_order[first_a_idx - 1]
                prev_sub = self.prodaja[(self.prodaja['id artikla']==ida) & (self.prodaja['Godina']==prev_god) & (self.prodaja['Mesec']==prev_mes)]
                start_lager = prev_sub['Lager'].sum() if len(prev_sub) > 0 else 0
                start_lager = start_lager if not pd.isna(start_lager) else 0
            lageri = [start_lager]
            for god, mes in a_meseci_order:
                msub = self.prodaja[(self.prodaja['id artikla']==ida) & (self.prodaja['Godina']==god) & (self.prodaja['Mesec']==mes)]
                lager_kraj = msub['Lager'].sum() if len(msub) > 0 else 0
                lageri.append(lager_kraj if not pd.isna(lager_kraj) else 0)
            avg_lager = np.mean(lageri)
            obrt = total_prod / avg_lager if avg_lager > 0 else 0
            dani_pokrivanja = (avg_lager / (total_prod / (n_a * 30))) if total_prod > 0 else 999
            n_obj_aktiv = sub[sub['Prodata Kolicina']>0]['ID KOMITENTA'].nunique()
            n_obj_total = sub['ID KOMITENTA'].nunique()
            prod_po_obj = total_prod / n_obj_aktiv if n_obj_aktiv > 0 else 0
            mes_prod = {}
            for _, r in sub.iterrows():
                key = f"{int(r['Godina'])}/{int(r['Mesec'])}"
                mes_prod[key] = mes_prod.get(key, 0) + int(r['Prodata Kolicina'])
            promo_rows.append({
                'id artikla': int(ida),
                'Naziv': sub.iloc[0]['Naziv artikla'],
                'Grupa': sub.iloc[0]['Grupa'],
                'Redovna': pi['redovna'], 'Akcijska': pi['akcijska'],
                'Popust_%': round((1 - pi['akcijska']/pi['redovna'])*100, 1),
                'Prodato_kom': total_prod,
                'Prihod_akcija': round(prihod_akcija, 0),
                'Prihod_redovna': round(prihod_redovna, 0),
                'Profit_akcija': round(profit_akcija, 0),
                'Profit_da_je_redovna': round(profit_redovna, 0),
                'Cena_akcije': round(razlika, 0),
                'Avg_lager': round(avg_lager, 0),
                'Obrt_x': round(obrt, 1),
                'Dani_pokrivanja': round(dani_pokrivanja, 0),
                'Obj_aktivnih': n_obj_aktiv,
                'Obj_ukupno': n_obj_total,
                'Prod_po_obj': round(prod_po_obj, 1),
                **{f'Prod_{a_labels[j]}': mes_prod.get(f"{int(a_meseci_order[j][0])}/{int(a_meseci_order[j][1])}", 0) for j in range(n_a)}
            })
        self.df_promo = pd.DataFrame(promo_rows).sort_values('Obrt_x', ascending=False)


def create_excel(engine):
    df=engine.df_result; ml=engine.mesec_labels; wb=Workbook()
    hf=PatternFill('solid',fgColor='2F5496'); hfn=Font(bold=True,color='FFFFFF',name='Arial',size=10)
    sfnt=Font(bold=True,name='Arial',size=9); dfn=Font(name='Arial',size=9)
    tb=Border(left=Side('thin','B4C6E7'),right=Side('thin','B4C6E7'),top=Side('thin','B4C6E7'),bottom=Side('thin','B4C6E7'))
    ca=Alignment(horizontal='center',vertical='center'); caw=Alignment(horizontal='center',vertical='center',wrap_text=True)
    sf_poc=PatternFill('solid',fgColor='D6E4F0'); sf_prom=PatternFill('solid',fgColor='C6EFCE')
    sf_prod=PatternFill('solid',fgColor='FFF2CC'); sf_pov=PatternFill('solid',fgColor='FCE4EC')
    sf_kor=PatternFill('solid',fgColor='E8E8E8'); sf_pred=PatternFill('solid',fgColor='D5A6E6')
    sf_avg=PatternFill('solid',fgColor='B4D7E8'); sf_razl=PatternFill('solid',fgColor='FFD699')
    sf_lager=PatternFill('solid',fgColor='DAEEF3'); sf_p1=PatternFill('solid',fgColor='92D050')
    sf_p2=PatternFill('solid',fgColor='00B050'); pred_hdr=PatternFill('solid',fgColor='7030A0')
    ord_hdr=PatternFill('solid',fgColor='375623'); sf_hist=PatternFill('solid',fgColor='E2D5F1')
    nf_money='#,##0'

    SC=5; sub_h=['Pocetno stanje','Promet (ulaz)','Prodaja','Povrat','Korekcija']
    sub_f=[sf_poc,sf_prom,sf_prod,sf_pov,sf_kor]; col_suf=['_Pocetno','_Promet','_Prodaja','_Povrat','_Korekcija']
    ws1=wb.active; ws1.title="Pregled po objektima"
    for c,t in enumerate(['ID Komitenta','ID Artikla','Naziv Artikla','Grupa'],1):
        cell=ws1.cell(1,c,t); cell.font=hfn; cell.fill=hf; cell.alignment=ca; cell.border=tb
        ws1.merge_cells(start_row=1,end_row=2,start_column=c,end_column=c)
    hist_col=5; month_start=5
    if engine.has_history:
        cell=ws1.cell(1,hist_col,'Jan-Avg 2025'); cell.font=hfn; cell.fill=PatternFill('solid',fgColor='6B3FA0')
        cell.alignment=ca; cell.border=tb
        ws1.merge_cells(start_row=1,end_row=1,start_column=hist_col,end_column=hist_col)
        c2=ws1.cell(2,hist_col,'Total prodaja'); c2.font=sfnt; c2.fill=sf_hist; c2.alignment=caw; c2.border=tb
        month_start=6
    for i,label in enumerate(ml):
        sc=month_start+i*SC
        ws1.merge_cells(start_row=1,end_row=1,start_column=sc,end_column=sc+SC-1)
        cell=ws1.cell(1,sc,label); cell.font=hfn; cell.fill=hf; cell.alignment=ca
        for cc in range(sc,sc+SC): ws1.cell(1,cc).border=tb; ws1.cell(1,cc).fill=hf
        for j,(sh,sfill) in enumerate(zip(sub_h,sub_f)):
            cell=ws1.cell(2,sc+j,sh); cell.font=sfnt; cell.fill=sfill; cell.border=tb; cell.alignment=caw
    ps=month_start+len(ml)*SC
    ws1.merge_cells(start_row=1,end_row=1,start_column=ps,end_column=ps+2)
    cell=ws1.cell(1,ps,f'{engine.pred_label} - PREDIKCIJA'); cell.font=hfn; cell.fill=pred_hdr; cell.alignment=ca
    for cc in range(ps,ps+3): ws1.cell(1,cc).border=tb; ws1.cell(1,cc).fill=pred_hdr
    for j,(sh,sfill) in enumerate(zip(['Predikcija','Prosek (svi mes.)','Razlika'],[sf_pred,sf_avg,sf_razl])):
        cell=ws1.cell(2,ps+j,sh); cell.font=sfnt; cell.fill=sfill; cell.border=tb; cell.alignment=caw
    os_c=ps+3
    ws1.merge_cells(start_row=1,end_row=1,start_column=os_c,end_column=os_c+2)
    cell=ws1.cell(1,os_c,f'PORUDZBINA - {engine.order_label}'); cell.font=hfn; cell.fill=ord_hdr; cell.alignment=ca
    for cc in range(os_c,os_c+3): ws1.cell(1,cc).border=tb; ws1.cell(1,cc).fill=ord_hdr
    ll="Lager danas"
    if len(engine.trenutni)>0:
        dc=[c for c in engine.trenutni.columns if 'dan' in c.lower()]
        if dc:
            try: d=pd.to_datetime(engine.trenutni[dc[0]].iloc[0]); ll=f"Lager na dan\n{d.strftime('%d.%m.%Y')}"
            except: pass
    for j,(sh,sfill) in enumerate(zip([ll,'Porudzbina\n(osnovna)',f'Porudzbina\n(min. {engine.min_lager} na stanju)'],[sf_lager,sf_p1,sf_p2])):
        cell=ws1.cell(2,os_c+j,sh); cell.font=sfnt; cell.fill=sfill; cell.border=tb; cell.alignment=caw
    for idx,row in df.iterrows():
        r=idx+3
        for c2,col in enumerate(['ID KOMITENTA','id artikla','Naziv artikla','Grupa'],1):
            ws1.cell(r,c2,row[col]).font=dfn; ws1.cell(r,c2).border=tb
        if engine.has_history:
            v=int(row.get('Total_JanAvg',0)); cell=ws1.cell(r,hist_col,v); cell.font=dfn
            cell.alignment=ca; cell.border=tb
            if v>0: cell.fill=PatternFill('solid',fgColor='F3EAFA')
        for i,label in enumerate(ml):
            cb=month_start+i*SC
            for j,suf in enumerate(col_suf):
                cn=f'{label}{suf}'; v=row.get(cn,0)
                cell=ws1.cell(r,cb+j,int(v) if not pd.isna(v) else 0); cell.font=dfn; cell.alignment=ca; cell.border=tb
        for j,cn in enumerate(['Predikcija','Prosek','Razlika']):
            v=int(row.get(cn,0)); cell=ws1.cell(r,ps+j,v); cell.alignment=ca; cell.border=tb
            if cn=='Razlika':
                if v>0: cell.font=Font(name='Arial',size=9,color='006100',bold=True)
                elif v<0: cell.font=Font(name='Arial',size=9,color='9C0006',bold=True)
                else: cell.font=dfn
            else: cell.font=dfn
        for j,cn in enumerate(['Lager_danas','Porudzbina_1','Porudzbina_2']):
            v=int(row.get(cn,0)); cell=ws1.cell(r,os_c+j,v); cell.alignment=ca; cell.border=tb
            if cn!='Lager_danas' and v>0: cell.font=Font(name='Arial',size=9,bold=True,color='375623')
            else: cell.font=dfn
    ws1.column_dimensions['A'].width=14; ws1.column_dimensions['B'].width=11; ws1.column_dimensions['C'].width=50; ws1.column_dimensions['D'].width=12
    if engine.has_history: ws1.column_dimensions[get_column_letter(hist_col)].width=14
    for i in range(len(ml)):
        for j in range(SC): ws1.column_dimensions[get_column_letter(month_start+i*SC+j)].width=14
    for j in range(3): ws1.column_dimensions[get_column_letter(ps+j)].width=14
    for j in range(3): ws1.column_dimensions[get_column_letter(os_c+j)].width=18
    ws1.freeze_panes=f'{get_column_letter(month_start)}3'
    ws1.auto_filter.ref=f"A2:{get_column_letter(ws1.max_column)}{ws1.max_row}"

    ws2=wb.create_sheet("Totali po mesecima")
    for c,h in enumerate(['Mesec','Promet (ulaz)','Prodaja','Stvarni povrat','Korekcija','Neto (Promet-Povrat)'],1):
        cell=ws2.cell(1,c,h); cell.font=hfn; cell.fill=hf; cell.alignment=caw; cell.border=tb
    ro=2
    if engine.has_history:
        ws2.cell(ro,1,'Jan-Avg 2025 (UKUPNO)').font=Font(bold=True,name='Arial',size=10,color='6B3FA0')
        ws2.cell(ro,1).alignment=ca; ws2.cell(ro,1).border=tb
        cell=ws2.cell(ro,3,int(df['Total_JanAvg'].sum())); cell.font=Font(bold=True,name='Arial',size=10,color='6B3FA0')
        cell.fill=sf_hist; cell.alignment=ca; cell.border=tb; cell.number_format=nf_money
        for c in [2,4,5,6]: ws2.cell(ro,c,'-').font=dfn; ws2.cell(ro,c).alignment=ca; ws2.cell(ro,c).border=tb
        ro+=2
    for ri,label in enumerate(ml,ro):
        ws2.cell(ri,1,label).font=Font(bold=True,name='Arial',size=10); ws2.cell(ri,1).alignment=ca; ws2.cell(ri,1).border=tb
        vals=[int(df[f'{label}_Promet'].sum()),int(df[f'{label}_Prodaja'].sum()),int(df[f'{label}_Povrat'].sum()),int(df[f'{label}_Korekcija'].sum())]
        vals.append(vals[0]-vals[2])
        fills=[sf_prom,sf_prod,sf_pov,sf_kor,sf_poc]
        for c2,(v,f) in enumerate(zip(vals,fills),2):
            cell=ws2.cell(ri,c2,v); cell.font=dfn; cell.fill=f; cell.alignment=ca; cell.border=tb; cell.number_format=nf_money
    fr=ro+len(ml)+1
    ws2.cell(fr,1,f'PORUDZBINA {engine.order_label.upper()}').font=Font(bold=True,name='Arial',size=11,color='375623'); ws2.cell(fr,1).border=tb
    ir=[(f'Predikcija {engine.pred_label}',int(df['Predikcija'].sum()),sf_pred),('Prosek (svi meseci)',int(df['Prosek'].sum()),sf_avg),
        ('Trenutni lager',int(df['Lager_danas'].sum()),sf_lager),
        ('Porudzbina (osnovna)',int(df[~df['ID KOMITENTA'].isin(engine.excluded)]['Porudzbina_1'].sum()),sf_p1),
        (f'Porudzbina (min. {engine.min_lager})',int(df[~df['ID KOMITENTA'].isin(engine.excluded)]['Porudzbina_2'].sum()),sf_p2)]
    for i,(label,val,fill) in enumerate(ir,fr+1):
        ws2.cell(i,1,label).font=Font(bold=True,name='Arial',size=10); ws2.cell(i,1).alignment=ca; ws2.cell(i,1).border=tb
        cell=ws2.cell(i,2,val); cell.font=Font(bold=True,name='Arial',size=11); cell.fill=fill; cell.alignment=ca; cell.border=tb; cell.number_format=nf_money
    ws2.column_dimensions['A'].width=32; ws2.column_dimensions['B'].width=18
    for c in 'CDEF': ws2.column_dimensions[c].width=18

    if engine.has_prices and len(engine.df_oos) > 0:
        ws_oos = wb.create_sheet("OOS Izgubljeni profit")
        oos_hdr = PatternFill('solid', fgColor='C00000')
        oos_fill = PatternFill('solid', fgColor='FCE4EC')
        a_labels_oos = engine.analitika_labels if engine.analitika_labels else engine.mesec_labels
        fixed_h = ['ID Komitenta','ID Artikla','Naziv','Grupa','Prosek kad ima','Lager danas']
        mes_h = []
        for lb in a_labels_oos: mes_h += [f'OOS {lb}', f'Izgub {lb} (RSD)']
        all_h = fixed_h + mes_h + ['OOS meseci ukupno','Izgubljeni profit (RSD)']
        for c, h in enumerate(all_h, 1):
            cell = ws_oos.cell(1, c, h)
            cell.font=Font(bold=True,color='FFFFFF',name='Arial',size=9)
            cell.fill=oos_hdr; cell.alignment=caw; cell.border=tb
        for idx, (_, row) in enumerate(engine.df_oos.iterrows(), 2):
            vals = [row['ID KOMITENTA'], row['id artikla'], row['Naziv artikla'], row['Grupa'],
                    row.get('Prosek_kad_ima',0), row.get('Lager_danas',0)]
            for lb in a_labels_oos:
                vals.append(int(row.get(f'OOS_{lb}', 0)))
                vals.append(row.get(f'Izgub_{lb}', 0))
            vals += [row.get('OOS_meseci',0), row.get('Izgubljeni_profit',0)]
            for c, v in enumerate(vals, 1):
                cell = ws_oos.cell(idx, c, v); cell.font=dfn; cell.border=tb; cell.alignment=ca
                col_name = all_h[c-1]
                if col_name.startswith('OOS ') and v == 1:
                    cell.fill = oos_fill; cell.font = Font(name='Arial',size=9,bold=True,color='C00000')
                if col_name.startswith('Izgub ') or col_name == 'Izgubljeni profit (RSD)':
                    cell.number_format = nf_money
                if col_name == 'Lager danas' and v == 0:
                    cell.fill = oos_fill; cell.font = Font(name='Arial',size=9,bold=True,color='C00000')
        ws_oos.column_dimensions['A'].width=13; ws_oos.column_dimensions['B'].width=10
        ws_oos.column_dimensions['C'].width=45; ws_oos.column_dimensions['D'].width=12
        ws_oos.column_dimensions['E'].width=14; ws_oos.column_dimensions['F'].width=12
        for i in range(len(a_labels_oos)*2):
            ws_oos.column_dimensions[get_column_letter(7+i)].width=13
        last_col = 7 + len(a_labels_oos)*2
        ws_oos.column_dimensions[get_column_letter(last_col)].width=14
        ws_oos.column_dimensions[get_column_letter(last_col+1)].width=18
        ws_oos.freeze_panes='E2'
        ws_oos.auto_filter.ref=f"A1:{get_column_letter(len(all_h))}{len(engine.df_oos)+1}"

    if engine.has_prices and len(engine.df_profit_obj) > 0:
        ws_prof = wb.create_sheet("Profitabilnost objekata")
        prof_hdr = PatternFill('solid', fgColor='1F4E79')
        bad_fill = PatternFill('solid', fgColor='FCE4EC')
        good_fill = PatternFill('solid', fgColor='E2EFDA')
        headers = ['ID Komitenta','Artikala','Prodato kom','Bruto profit (RSD)','Trosak mkt (RSD)','Neto profit (RSD)','Izgubljeno OOS (RSD)','Potencijal (RSD)']
        for lb in (engine.analitika_labels if engine.analitika_labels else ml): headers.append(f'Neto {lb}')
        for c, h in enumerate(headers, 1):
            cell = ws_prof.cell(1, c, h); cell.font=Font(bold=True,color='FFFFFF',name='Arial',size=9); cell.fill=prof_hdr; cell.alignment=caw; cell.border=tb
        for idx, (_, row) in enumerate(engine.df_profit_obj.iterrows(), 2):
            vals = [row['ID KOMITENTA'], row['Artikala'], row['Prodato_kom'], row['Bruto_profit'],
                    row['Trosak_mkt'], row['Neto_profit'], row['Izgubljeno_OOS'], row['Potencijalni_profit']]
            for lb in (engine.analitika_labels if engine.analitika_labels else ml): vals.append(row.get(f'Neto_{lb}', 0))
            for c, v in enumerate(vals, 1):
                cell = ws_prof.cell(idx, c, v); cell.font=dfn; cell.border=tb; cell.alignment=ca
                if c >= 4: cell.number_format=nf_money
                if c == 6:
                    if v <= 0: cell.fill = bad_fill; cell.font = Font(name='Arial', size=9, bold=True, color='C00000')
                    elif v > 0: cell.fill = good_fill
                if c >= 9:
                    if v < 0: cell.font = Font(name='Arial', size=9, color='C00000')
                    elif v > 0: cell.font = Font(name='Arial', size=9, color='006100')
        for cl in 'AB': ws_prof.column_dimensions[cl].width=13
        ws_prof.column_dimensions['C'].width=12
        for cl in 'DEFGH': ws_prof.column_dimensions[cl].width=18
        a_ml = engine.analitika_labels if engine.analitika_labels else ml
        for i in range(len(a_ml)): ws_prof.column_dimensions[get_column_letter(9+i)].width=14
        ws_prof.freeze_panes='B2'
        ws_prof.auto_filter.ref=f"A1:{get_column_letter(len(headers))}{len(engine.df_profit_obj)+1}"

    if engine.has_prices and len(engine.df_promo) > 0:
        ws_akc = wb.create_sheet("Analiza akcije")
        akc_hdr = PatternFill('solid', fgColor='BF8F00')
        good_obrt = PatternFill('solid', fgColor='E2EFDA')
        bad_obrt = PatternFill('solid', fgColor='FCE4EC')
        headers = ['ID Artikla','Naziv','Grupa','Redovna\ncena','Akcijska\ncena','Popust\n%',
                   'Prodato\nkom','Prihod\nakcija (RSD)','Prihod da je\nredovna (RSD)',
                   'Profit\nakcija (RSD)','Profit da je\nredovna (RSD)','Cena akcije\n(RSD)',
                   'Prosecni\nlager','Obrt\n(x)','Dani\npokrivanja',
                   'Aktivnih\nobjekata','Ukupno\nobjekata','Prod.\npo objektu']
        for lb in (engine.analitika_labels if engine.analitika_labels else ml): headers.append(f'Prod.\n{lb}')
        for c, h in enumerate(headers, 1):
            cell = ws_akc.cell(1, c, h); cell.font=Font(bold=True,color='FFFFFF',name='Arial',size=9); cell.fill=akc_hdr; cell.alignment=caw; cell.border=tb
        for idx, (_, row) in enumerate(engine.df_promo.iterrows(), 2):
            vals = [row['id artikla'], row['Naziv'], row['Grupa'], row['Redovna'], row['Akcijska'],
                    row['Popust_%'], row['Prodato_kom'],
                    row['Prihod_akcija'], row['Prihod_redovna'],
                    row['Profit_akcija'], row['Profit_da_je_redovna'], row['Cena_akcije'],
                    row['Avg_lager'], row['Obrt_x'], row['Dani_pokrivanja'],
                    row['Obj_aktivnih'], row['Obj_ukupno'], row['Prod_po_obj']]
            for lb in (engine.analitika_labels if engine.analitika_labels else ml): vals.append(row.get(f'Prod_{lb}', 0))
            for c, v in enumerate(vals, 1):
                cell = ws_akc.cell(idx, c, v); cell.font=dfn; cell.border=tb; cell.alignment=ca
                if c in [4,5,8,9,10,11,12]: cell.number_format=nf_money
                if c == 14:
                    if v >= 2.0: cell.fill = good_obrt; cell.font = Font(name='Arial',size=9,bold=True,color='006100')
                    elif v < 1.0: cell.fill = bad_obrt; cell.font = Font(name='Arial',size=9,bold=True,color='C00000')
                if c == 15 and v > 120: cell.fill = bad_obrt
        ws_akc.column_dimensions['A'].width=10; ws_akc.column_dimensions['B'].width=45; ws_akc.column_dimensions['C'].width=12
        for cl in 'DEFG': ws_akc.column_dimensions[cl].width=12
        for cl in 'HIJKL': ws_akc.column_dimensions[cl].width=16
        for cl in 'MNOPQR': ws_akc.column_dimensions[cl].width=13
        a_ml2 = engine.analitika_labels if engine.analitika_labels else ml
        for i in range(len(a_ml2)): ws_akc.column_dimensions[get_column_letter(19+i)].width=11
        ws_akc.auto_filter.ref=f"A1:{get_column_letter(len(headers))}{len(engine.df_promo)+1}"

    ws3=wb.create_sheet("O modelu"); ws3.column_dimensions['A'].width=100
    info=["OPIS MODELA PREDIKCIJE I PORUDZBINE","",f"=== PREDIKCIJA ZA {engine.pred_label.upper()} ===","",
        "Model predvidja POTENCIJAL PRODAJE.","",
        f"  1. Constrained sales korekcija:",
        f"     - Kraj meseca lager=0 i prodaja>0: rasprodato, potraznja veca — zameni prosekom normalnih meseci",
        f"     - Pocetno=0 i promet=0: cist OOS — zameni prosekom normalnih meseci",
        f"     - Pocetno=0 i promet>0 i kraj=0: dobili i rasprodali — zameni prosekom",
        f"     - Normalni meseci = ostalo robe na kraju (lager>0)",
        f"  2. Holt DES (alpha={engine.alpha}, beta={engine.beta}) + WMA (50/28/12/7/3%)",
        "  3. Kombinacija: 60% veci + 40% manji od Holt/WMA",
        "  4. Varijansa boost (faktor 0.4, max 70%)",
        "  5. Niska zaliha (0-2): predikcija minimum prosek kad je na stanju",
        "  6. Prodaja 5+ mesecno: predikcija minimum prosek",
        "  7. Donje ogranicenje: predikcija < prosek samo ako poslednjih 5 meseci pada ili stagnira (<=)",
        "  8. Sigurnosna mreza: predikcija=0 samo ako nista prodato u poslednjih 5 meseci; ako poslednji mesec >1 onda min taj broj",
        "  9. Zaokruzivanje: round (predikcija i prosek)",
        ]
    if engine.has_history: info+=[f"  10. Istorijski podaci: {HIST_WEIGHT*100:.0f}% tezina"]
    info+=["",f"=== PORUDZBINA ZA {engine.order_label.upper()} ===","",
        f"P1 (osnovna): max(Pred-Lager, 0)",
        f"P2 (sa dopunom): Za lager<=2: dopuna do max(predikcija, prosek, min porudzbina={engine.min_order}); Za lager>2: dopuna do min {engine.min_lager}",
        f"P2 finalna provera: ako (P2+lager) <= prodaja_poslednjeg_meseca, dodaje se buffer (1-5 kom: +2, 6-10: +3, 11-15: +4, 16+: +5)",
        f"Iskljuceni: {', '.join(str(x) for x in sorted(engine.excluded))}"]
    if engine.has_prices:
        info+=["",f"=== ANALITIKA ===","",
            f"Profit formula: (Finalna cena / 1.2 / 1.2 - Nabavna) x Kolicina",
            f"OOS izgubljeni profit: prosek prodaje kad ima zaliha x OOS meseci x profit/kom",
            f"Ukupan trosak marketinga: {engine.mesecni_trosak:,.0f} RSD / {engine.num_komitenti} objekata = {engine.trosak_po_objektu:,.0f} RSD po objektu za period",
            f"Mesecni trosak po objektu: {engine.trosak_po_objektu / max(len(engine.analitika_labels), 1):,.0f} RSD",
            f"Neto po mesecu = Bruto profit meseca - mesecni trosak po objektu"]
    info+=[f"","Generisano: {datetime.datetime.now().strftime('%d.%m.%Y. u %H:%M')}"]
    for i,line in enumerate(info,1):
        cell=ws3.cell(i,1,line)
        if i==1: cell.font=Font(bold=True,name='Arial',size=14,color='375623')
        elif '===' in line: cell.font=Font(bold=True,name='Arial',size=12,color='7030A0')
        else: cell.font=Font(name='Arial',size=10)

    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf


DEFAULT_EXCLUDED = "1023, 1027, 1034, 1043, 1057, 1060, 1061, 1076, 1315, 1347, 1349, 1359"

st.set_page_config(page_title="VAPE Analitika", page_icon="\U0001f4a8", layout="wide", initial_sidebar_state="collapsed")

# --- Sidebar: hidden on report pages ---
_pg = st.session_state.get('page', 'home')
if _pg not in ('home', 'porudzbine'):
    st.markdown("""<style>
    section[data-testid="stSidebar"] {
        width: 0px !important;
        min-width: 0px !important;
        overflow: hidden !important;
        visibility: hidden !important;
    }
    .main .block-container {
        padding-left: 2rem !important;
        padding-right: 2rem !important;
        max-width: 100% !important;
    }
    </style>""", unsafe_allow_html=True)


# =====================================================================
# NOVI DIZAJN — samo boje i layout, nista matematicko se ne menja
# =====================================================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap');

    .stApp {
        background: #f5f0ff !important;
        font-family: 'Poppins', sans-serif;
    }

    /* --- SIDEBAR --- */
    section[data-testid="stSidebar"] {
        background: #12002a !important;
        border-right: 1px solid rgba(168,85,247,0.15) !important;
    }
    section[data-testid="stSidebar"] * { color: rgba(255,255,255,0.85) !important; }
    section[data-testid="stSidebar"] h3 {
        color: white !important;
        font-size: 13px !important;
        font-weight: 700 !important;
        letter-spacing: 0.5px !important;
        text-transform: uppercase !important;
        margin-bottom: 8px !important;
    }
    section[data-testid="stSidebar"] input,
    section[data-testid="stSidebar"] textarea {
        background: rgba(255,255,255,0.08) !important;
        border: 1px solid rgba(255,255,255,0.15) !important;
        color: white !important;
        border-radius: 8px !important;
    }
    section[data-testid="stSidebar"] input:focus,
    section[data-testid="stSidebar"] textarea:focus {
        border-color: rgba(168,85,247,0.5) !important;
        box-shadow: 0 0 0 2px rgba(168,85,247,0.15) !important;
    }
    section[data-testid="stSidebar"] hr {
        border-color: rgba(255,255,255,0.1) !important;
    }

    /* --- INPUTI VAN SIDEBARA (glavni sadrzaj) --- */
    .main .stTextInput > div > div > input,
    .main .stNumberInput > div > div > input {
        background: white !important;
        border: 1px solid rgba(168,85,247,0.25) !important;
        color: #1a0533 !important;
        border-radius: 8px !important;
    }
    .main .stTextInput > div > div > input::placeholder,
    .main .stNumberInput > div > div > input::placeholder {
        color: #9ca3af !important;
    }
    .main .stTextInput > div > div > input:focus,
    .main .stNumberInput > div > div > input:focus {
        border-color: #a855f7 !important;
        box-shadow: 0 0 0 2px rgba(168,85,247,0.15) !important;
    }

    /* Sidebar logo traka */
    section[data-testid="stSidebar"]::before {
        content: '';
        display: block;
        height: 4px;
        background: linear-gradient(90deg, #a855f7, #ec4899);
        margin-bottom: 0;
    }

    /* --- METRIC KARTICE --- */
    .metric-card {
        background: white;
        border-radius: 14px;
        padding: 16px 20px;
        box-shadow: 0 2px 12px rgba(124,58,237,0.07);
        border: 1px solid rgba(168,85,247,0.12);
        text-align: center;
    }
    .metric-value {
        font-size: 26px; font-weight: 700;
        background: linear-gradient(135deg, #7c3aed, #ec4899);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    }
    .metric-value-red { font-size: 26px; font-weight: 700; color: #dc2626; }
    .metric-value-green { font-size: 26px; font-weight: 700; color: #059669; }
    .metric-label { font-size: 11px; color: #888; margin-top: 4px; }

    /* --- DUGMAD --- */
    .stButton > button {
        background: linear-gradient(135deg, #a855f7 0%, #ec4899 100%) !important;
        color: white !important;
        border: none !important;
        border-radius: 12px !important;
        padding: 14px 32px !important;
        font-weight: 700 !important;
        font-size: 15px !important;
        box-shadow: 0 4px 15px rgba(168,85,247,0.3) !important;
        transition: opacity 0.2s !important;
    }
    .stButton > button:hover { opacity: 0.88 !important; }

    /* --- DOWNLOAD DUGME --- */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important;
        color: white !important;
        border: none !important;
        border-radius: 12px !important;
        padding: 14px 32px !important;
        font-weight: 700 !important;
        box-shadow: 0 4px 15px rgba(16,185,129,0.25) !important;
    }

    /* --- MULTISELECT TAGOVI --- */
    .stMultiSelect [data-baseweb="tag"] {
        background: linear-gradient(135deg, #a855f7, #ec4899) !important;
        border: none !important;
        border-radius: 99px !important;
        color: white !important;
        font-weight: 600 !important;
        font-size: 12px !important;
    }
    .stMultiSelect [data-baseweb="tag"] span { color: white !important; }
    .stMultiSelect [data-baseweb="tag"] button { color: rgba(255,255,255,0.8) !important; }
    .stMultiSelect [data-baseweb="select"] > div {
        border: 1px solid rgba(168,85,247,0.3) !important;
        border-radius: 10px !important;
        background: white !important;
    }
    .stMultiSelect [data-baseweb="select"] > div:focus-within {
        border-color: #a855f7 !important;
        box-shadow: 0 0 0 2px rgba(168,85,247,0.15) !important;
    }

    /* --- HEADER TRAKA (bez kvadrata) --- */
    .header-navbar {
        background: #12002a;
        border-radius: 0;
        padding: 0 32px;
        height: 56px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin: -1rem -1rem 24px -1rem;
        border-bottom: 3px solid transparent;
        border-image: linear-gradient(90deg, #a855f7, #ec4899) 1;
    }

    /* --- SUCCESS / WARN BOXOVI --- */
    .success-box {
        background: linear-gradient(135deg, rgba(16,185,129,0.08), rgba(5,150,105,0.04));
        border: 1px solid rgba(16,185,129,0.2);
        border-radius: 10px;
        padding: 12px 16px;
    }
    .warn-box {
        background: linear-gradient(135deg, rgba(220,38,38,0.07), rgba(220,38,38,0.02));
        border: 1px solid rgba(220,38,38,0.18);
        border-radius: 10px;
        padding: 10px 14px;
        margin: 6px 0;
    }
    .section-title {
        font-size: 17px; font-weight: 600; color: #4c1d95; margin: 16px 0 8px 0;
    }
</style>
""", unsafe_allow_html=True)

# === SESSION STATE NAVIGACIJA ===
if 'page' not in st.session_state:
    st.session_state.page = 'home'

alpha = 0.4
beta = 0.2

# === SIDEBAR ===
with st.sidebar:
    st.markdown("""
    <div style="padding: 16px 4px 8px 4px; display:flex; align-items:center; gap:10px; margin-bottom:4px;">
        <div style="width:26px; height:26px; background:linear-gradient(135deg,#a855f7,#ec4899);
            border-radius:7px; display:flex; align-items:center; justify-content:center; flex-shrink:0;">
            <div style="width:9px; height:9px; background:white; border-radius:2px;"></div>
        </div>
        <span style="font-size:15px; font-weight:700; color:white;">VAPE</span>
        <span style="font-size:15px; font-weight:300; color:rgba(255,255,255,0.4);">Analitika</span>
    </div>
    <div style="height:1px; background:linear-gradient(90deg,rgba(168,85,247,0.5),rgba(236,72,153,0.3),transparent); margin-bottom:16px;"></div>
    """, unsafe_allow_html=True)

    # Navigaciona dugmad
    st.markdown("##### Navigacija")
    if st.button("🏠  Početna", use_container_width=True):
        st.session_state.page = 'home'
        st.rerun()
    if st.button("📦  Profitabilnost objekata", use_container_width=True):
        st.session_state.page = 'porudzbine'
        st.rerun()
    if st.button("📊  Mesečni izveštaj prodaje", use_container_width=True):
        st.session_state.page = 'mesecni'
        st.rerun()
    if st.button("💰  Finansijski izveštaj", use_container_width=True):
        st.session_state.page = 'finansijski'
        st.rerun()

    st.markdown("---")

    # Parametri samo za Profitabilnost
    if st.session_state.page == 'porudzbine':
        st.markdown("### 📦 Parametri porudžbine")
        _ml_str = st.text_input("Minimalni lager po artiklu", value="", placeholder="prazno = bez ograničenja",
            help="Dopuni objekat da ima minimum X komada na stanju po artiklu. Ostavi prazno za bez ograničenja.")
        min_lager = int(_ml_str) if _ml_str.strip().isdigit() else None
        _mo_str = st.text_input("Min. ukupna porudžbina po objektu", value="", placeholder="prazno = bez ograničenja",
            help="Ako je ukupna porudžbina za objekat manja od X, ne šalji ništa. Ostavi prazno za bez ograničenja.")
        min_order = int(_mo_str) if _mo_str.strip().isdigit() else None
        st.markdown("---")
        st.markdown("### 💰 Troškovi")
        mesecni_trosak = st.number_input(
            "Ukupan trosak mkt/ulistavanja za ceo period (RSD)",
            min_value=0, value=0, step=10000,
            help="Unesi UKUPAN iznos za ceo analizirani period — automatski se deli na broj objekata i broj meseci"
        )
        st.markdown("---")
        st.markdown("### ⛔ Isključeni komitenti")
        excluded_str = st.text_area("ID-evi razdvojeni zarezom", value=DEFAULT_EXCLUDED, height=100)
    else:
        min_lager = None
        min_order = None
        mesecni_trosak = 0
        excluded_str = DEFAULT_EXCLUDED

excluded = set()
for part in excluded_str.replace('\n', ',').split(','):
    p = part.strip()
    if p.isdigit(): excluded.add(int(p))

# === HEADER FUNCTION ===
def render_header(subtitle):
    st.markdown(f'''<div style="background:#12002a;border-radius:16px;padding:0 28px;height:60px;
        display:flex;align-items:center;justify-content:space-between;margin-bottom:24px;
        border-bottom:3px solid;border-image:linear-gradient(90deg,#a855f7,#ec4899) 1;
        box-shadow:0 4px 20px rgba(18,0,42,0.18);">
        <div style="display:flex;align-items:center;gap:12px;">
            <div style="width:30px;height:30px;background:linear-gradient(135deg,#a855f7,#ec4899);
                border-radius:8px;display:flex;align-items:center;justify-content:center;">
                <div style="width:11px;height:11px;background:white;border-radius:3px;"></div>
            </div>
            <span style="font-size:18px;font-weight:700;color:white;">VAPE</span>
            <span style="font-size:18px;font-weight:300;color:rgba(255,255,255,0.4);">Analitika</span>
            <span style="font-size:11px;color:rgba(255,255,255,0.25);margin-left:8px;">·</span>
            <span style="font-size:12px;color:rgba(255,255,255,0.35);">{subtitle}</span>
        </div>
        <div style="display:flex;gap:6px;align-items:center;">
            <div style="width:8px;height:8px;border-radius:50%;background:rgba(168,85,247,0.7);"></div>
            <div style="width:8px;height:8px;border-radius:50%;background:rgba(236,72,153,0.5);"></div>
            <div style="width:8px;height:8px;border-radius:50%;background:rgba(255,255,255,0.15);"></div>
        </div>
    </div>''', unsafe_allow_html=True)
    if st.session_state.get('page', 'home') != 'home':
        st.markdown("""<script>
        (function() {
            function collapseSidebar() {
                var btn = window.parent.document.querySelector('[data-testid="collapsedControl"]');
                var sidebar = window.parent.document.querySelector('[data-testid="stSidebar"]');
                if (sidebar) {
                    var expanded = sidebar.getAttribute('aria-expanded');
                    if (expanded === 'false') return;
                }
                if (btn) { setTimeout(function(){ btn.click(); }, 300); }
            }
            if (document.readyState === 'complete') collapseSidebar();
            else window.addEventListener('load', collapseSidebar);
        })();
        </script>""", unsafe_allow_html=True)

page = st.session_state.page

# ============================================================
# LANDING PAGE
# ============================================================
if page == 'home':
    render_header("Odaberi izveštaj")

    components.html("""<!DOCTYPE html><html><head>
<link href="https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;600;700&display=swap" rel="stylesheet">
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Poppins',sans-serif;background:transparent;padding:24px 16px}
.label{font-size:10px;font-weight:600;letter-spacing:1px;text-transform:uppercase;color:#9ca3af;margin-bottom:20px;text-align:center}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:18px;max-width:820px;margin:0 auto}
.card{background:white;border-radius:20px;padding:24px 20px 20px;border:2px solid transparent;
    box-shadow:0 4px 20px rgba(124,58,237,0.08);transition:all 0.22s}
.card:hover{border-color:#a855f7;transform:translateY(-3px);box-shadow:0 12px 32px rgba(168,85,247,0.15)}
.icon{font-size:30px;margin-bottom:12px}
.title{font-size:15px;font-weight:700;color:#1a0533;margin-bottom:6px}
.desc{font-size:11px;color:#6b7280;line-height:1.6;margin-bottom:12px}
.tag{display:inline-block;font-size:10px;padding:2px 9px;border-radius:99px;font-weight:600}
.tag-purple{background:rgba(168,85,247,0.1);color:#7c3aed}
.tag-pink{background:rgba(236,72,153,0.09);color:#be185d}
.tag-blue{background:rgba(59,130,246,0.09);color:#1d4ed8}
</style></head><body>
<p class="label">AMAN d.o.o. · Odaberi izveštaj</p>
<div class="grid">
  <div class="card">
    <div class="icon">📦</div>
    <div class="title">Profitabilnost objekata</div>
    <div class="desc">Predikcija prodaje, OOS analiza, trendovi komitenata i analiza akcije.</div>
    <span class="tag tag-purple">Upload Excel</span>
  </div>
  <div class="card">
    <div class="icon">📊</div>
    <div class="title">Mesečni izveštaj</div>
    <div class="desc">Prodaja po sistemima, profitabilnost, Dr Vukašin i stanje zaliha.</div>
    <span class="tag tag-pink">Automatski podaci</span>
  </div>
  <div class="card">
    <div class="icon">💰</div>
    <div class="title">Finansijski izveštaj</div>
    <div class="desc">Pregled po sistemima, dugovanja, lager vrednosti i PDF generisanje.</div>
    <span class="tag tag-blue">Automatski podaci</span>
  </div>
</div>
</body></html>""", height=300)

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("📦 Otvori izveštaj", use_container_width=True, key='btn_home_p'):
            st.session_state.page = 'porudzbine'; st.rerun()
    with col2:
        if st.button("📊 Otvori izveštaj", use_container_width=True, key='btn_home_m'):
            st.session_state.page = 'mesecni'; st.rerun()
    with col3:
        if st.button("💰 Otvori izveštaj", use_container_width=True, key='btn_home_f'):
            st.session_state.page = 'finansijski'; st.rerun()

elif page == 'porudzbine':
    render_header("Predikcija prodaje · Profitabilnost · OOS analiza · Efekti akcije")
    uploaded = st.file_uploader("Učitaj Excel fajl sa podacima", type=['xlsx','xls'])

    if uploaded:
        file_bytes = uploaded.read()
        st.markdown(f'<div class="success-box">✅ Fajl <strong>{uploaded.name}</strong> učitan ({len(file_bytes)//1024} KB)</div>', unsafe_allow_html=True)
        st.markdown("")
        try:
            _xls = pd.ExcelFile(io.BytesIO(file_bytes))
            _sm = {s.strip().lower(): s for s in _xls.sheet_names}
            _sp = None
            for kw in ['prodaja']:
                for nl, no in _sm.items():
                    if kw in nl: _sp = no; break
            if _sp:
                _prod = pd.read_excel(_xls, sheet_name=_sp); _prod.columns=[c.strip() for c in _prod.columns]
                _meseci = sorted(_prod[['Godina','Mesec']].drop_duplicates().values.tolist())
                _mn={1:'Jan',2:'Feb',3:'Mar',4:'Apr',5:'Maj',6:'Jun',7:'Jul',8:'Avg',9:'Sep',10:'Okt',11:'Nov',12:'Dec'}
                _labels = [f"{_mn.get(int(m),'?' )} {int(g)}" for g,m in _meseci]
                st.markdown('**📅 Period za analizu** (OOS, Profitabilnost, Akcija — ne utiče na predikciju):')
                selected_labels = st.multiselect("Odaberi mesece", _labels, default=_labels, help="Predikcija uvek koristi sve mesece. Ovaj filter se odnosi samo na analitiku.")
                if not selected_labels:
                    st.warning("⚠️ Mora biti odabran bar jedan mesec za analizu. Automatski je odabran poslednji mesec.")
                    selected_labels = [_labels[-1]] if _labels else []
                selected_meseci = [_meseci[i] for i, lb in enumerate(_labels) if lb in selected_labels]
            else:
                selected_labels = []; selected_meseci = []
        except:
            selected_labels = []; selected_meseci = []

        if st.button("🚀 POKRENI ANALIZU", use_container_width=True):
            progress_bar = st.progress(0)
            try:
                engine = PredictionEngine(file_bytes, excluded, alpha, beta, min_lager, min_order, mesecni_trosak, selected_meseci)
                result = engine.run(progress_bar)

                st.markdown("---")
                tp = int(result['Predikcija'].sum()); tl = int(result['Lager_danas'].sum())
                t1 = int(result[~result['ID KOMITENTA'].isin(excluded)]['Porudzbina_1'].sum())
                t2 = int(result[~result['ID KOMITENTA'].isin(excluded)]['Porudzbina_2'].sum())

                if engine.has_prices:
                    tab1, tab2 = st.tabs(["📦 Porudžbina", "💰 Profitabilnost objekata & OOS"])
                else:
                    tab1, = st.tabs(["📦 Porudžbina"])

                with tab1:
                    n_obj_salji = int(result[result['Porudzbina_2'] > 0]['ID KOMITENTA'].nunique())
                    tp_prosek = int(result['Prosek'].sum())
                    m1,m2,m3,m4,m5 = st.columns(5)
                    m1.markdown(f'<div class="metric-card"><div class="metric-value">{tp:,}</div><div class="metric-label">Predikcija (kom)</div></div>', unsafe_allow_html=True)
                    m2.markdown(f'<div class="metric-card"><div class="metric-value">{tp_prosek:,}</div><div class="metric-label">Prosek (kom)</div></div>', unsafe_allow_html=True)
                    m3.markdown(f'<div class="metric-card"><div class="metric-value-green">{t2:,}</div><div class="metric-label">Porudžbina (kom)</div></div>', unsafe_allow_html=True)
                    m4.markdown(f'<div class="metric-card"><div class="metric-value">{n_obj_salji:,}</div><div class="metric-label">Objekata prima robu</div></div>', unsafe_allow_html=True)
                    m5.markdown(f'<div class="metric-card"><div class="metric-value">{tl:,}</div><div class="metric-label">Lager danas</div></div>', unsafe_allow_html=True)
                    st.markdown("")

                    st.markdown("<div style='margin:24px 0 4px 0;'></div>", unsafe_allow_html=True)

                    ml = engine.mesec_labels
                    df_r = engine.df_result.copy()

                    kom_mes = {}
                    for lb in ml:
                        col_lb = f'{lb}_Prodaja'
                        if col_lb in df_r.columns:
                            grp = df_r.groupby('ID KOMITENTA')[col_lb].sum()
                            for kid, v in grp.items():
                                if kid not in kom_mes: kom_mes[kid] = {}
                                kom_mes[kid][lb] = int(v)

                    import numpy as _np2

                    def _is_rastuci(vals5, dozvoljeni_sum=1):
                        padovi = sum(1 for i in range(1, len(vals5)) if vals5[i] < vals5[i-1])
                        return padovi <= dozvoljeni_sum and vals5[-1] > vals5[0] and sum(vals5) >= 10

                    def _is_padajuci(vals5, dozvoljeni_sum=1):
                        rasti = sum(1 for i in range(1, len(vals5)) if vals5[i] > vals5[i-1])
                        return rasti <= dozvoljeni_sum and vals5[-1] < vals5[0] and sum(vals5) >= 10

                    def _rast_pct(vals5):
                        first = vals5[0] if vals5[0] > 0 else 1
                        return (vals5[-1] - vals5[0]) / first * 100

                    rastuci_list = []
                    padajuci_list = []

                    for kid, mes_vals in kom_mes.items():
                        vals_all = [mes_vals.get(lb, 0) for lb in ml]
                        vals5 = vals_all[-5:] if len(vals_all) >= 5 else vals_all
                        if len(vals5) < 3: continue
                        if _is_rastuci(vals5):
                            rastuci_list.append({
                                'ID': kid, 'Ukupno': sum(vals_all),
                                'Vals': vals_all, 'Vals5': vals5,
                                'Rast': _rast_pct(vals5),
                                'Zadnji': vals5[-1], 'Prvi': vals5[0],
                            })
                        elif _is_padajuci(vals5):
                            padajuci_list.append({
                                'ID': kid, 'Ukupno': sum(vals_all),
                                'Vals': vals_all, 'Vals5': vals5,
                                'Pad': _rast_pct(vals5),
                                'Zadnji': vals5[-1], 'Prvi': vals5[0],
                            })

                    rastuci_list = sorted(rastuci_list, key=lambda x: x['Rast'], reverse=True)[:10]
                    padajuci_list = sorted(padajuci_list, key=lambda x: x['Pad'])[:10]

                    def _render_trend_section(title, icon, color, items, is_rast):
                        label_color = "#10b981" if is_rast else "#ef4444"
                        label_bg = "#f0fdf4" if is_rast else "#fef2f2"
                        if not items:
                            components.html(f"""<!DOCTYPE html><html><body style="margin:0;padding:4px 0;font-family:'DM Sans',sans-serif;">
                            <div style="display:flex;align-items:center;gap:8px;margin-bottom:12px;">
                                <span style="font-size:17px;">{icon}</span>
                                <span style="font-size:13px;font-weight:700;color:#111;">{title}</span>
                            </div>
                            <div style="color:#aaa;font-size:13px;padding:12px 0;">Nema podataka za prikaz</div>
                            </body></html>""", height=80)
                            return
                        rows_html = ""
                        for r in items:
                            vals5 = r['Vals5']
                            mx = max(vals5) if max(vals5) > 0 else 1
                            bars = "".join(
                                f'<div style="flex:1;display:flex;flex-direction:column;justify-content:flex-end;gap:0;">'
                                f'<div style="height:{int(v/mx*28)}px;background:{"linear-gradient(180deg,#a855f7,#c084fc)" if is_rast else "linear-gradient(180deg,#ec4899,#f9a8d4)"};border-radius:2px 2px 0 0;min-height:2px;"></div></div>'
                                for v in vals5
                            )
                            sign = "+" if is_rast else ""
                            pct = r['Rast'] if is_rast else r['Pad']
                            rows_html += f"""<div style="display:flex;align-items:center;gap:10px;padding:8px 0;border-bottom:1px solid #f3f4f6;">
                                <div style="font-family:'DM Mono',monospace;font-size:14px;font-weight:500;color:#111;width:46px;flex-shrink:0;">{int(r["ID"])}</div>
                                <div style="display:flex;align-items:flex-end;gap:2px;height:32px;width:90px;flex-shrink:0;">{bars}</div>
                                <div style="flex:1;font-size:11px;color:#aaa;">{int(r["Ukupno"]):,} kom</div>
                                <div style="font-size:12px;font-weight:700;color:{label_color};white-space:nowrap;">{sign}{pct:.0f}% &nbsp;<span style="font-weight:400;color:#bbb;font-size:11px;">({int(r["Prvi"])}→{int(r["Zadnji"])})</span></div>
                            </div>"""
                        h_px = len(items) * 48 + 56
                        components.html(f"""<!DOCTYPE html><html>
                        <head><link href="https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=DM+Sans:wght@400;600;700&display=swap" rel="stylesheet"></head>
                        <body style="margin:0;padding:4px 0;font-family:'DM Sans',sans-serif;background:white;">
                            <div style="display:flex;align-items:center;gap:8px;margin-bottom:14px;">
                                <span style="font-size:17px;">{icon}</span>
                                <span style="font-size:13px;font-weight:700;color:#111;">{title}</span>
                                <span style="font-size:10px;font-weight:700;color:{label_color};background:{label_bg};border-radius:20px;padding:2px 8px;">zadnjih 5 mes.</span>
                            </div>
                            <div style="font-size:9px;color:#ccc;display:flex;gap:10px;margin-bottom:4px;">
                                <span style="width:46px;"></span>
                                <span style="width:90px;text-align:center;text-transform:uppercase;letter-spacing:.5px;">trend</span>
                                <span style="flex:1;text-transform:uppercase;letter-spacing:.5px;">ukupno</span>
                                <span style="text-transform:uppercase;letter-spacing:.5px;">rast (prvi→zadnji)</span>
                            </div>
                            {rows_html}
                        </body></html>""", height=h_px)

                    def _render_oos_section(items, max_val):
                        if not items:
                            components.html('''<!DOCTYPE html><html><body style="margin:0;padding:4px 0;font-family:sans-serif;">
                            <div style="display:flex;align-items:center;gap:8px;margin-bottom:12px;">
                                <span style="font-size:17px;">🔴</span>
                                <span style="font-size:13px;font-weight:700;color:#111;">OOS — Lager 0, najveći potencijal</span>
                            </div>
                            <div style="color:#aaa;font-size:13px;">Nema OOS podataka</div>
                            </body></html>''', height=80)
                            return
                        rows_html = ""
                        for r in items:
                            pct = int(r['Izgubljeno'] / max_val * 100)
                            rows_html += f"""<div style="padding:9px 0;border-bottom:1px solid #f9f9f9;">
                                <div style="display:flex;align-items:center;gap:10px;margin-bottom:5px;">
                                    <div style="font-family:'DM Mono',monospace;font-size:14px;font-weight:500;color:#111;width:46px;flex-shrink:0;">{int(r["ID KOMITENTA"])}</div>
                                    <div style="font-size:10px;font-weight:700;color:#ec4899;background:#fdf2f8;border-radius:4px;padding:2px 7px;">{int(r["Artikala"])} artikala bez robe</div>
                                    <div style="margin-left:auto;font-family:'DM Mono',monospace;font-size:13px;font-weight:700;color:#7c3aed;">{int(r["Izgubljeno"]):,} RSD</div>
                                </div>
                                <div style="height:5px;background:#f5f0ff;border-radius:99px;overflow:hidden;">
                                    <div style="width:{pct}%;height:100%;background:linear-gradient(90deg,#a855f7,#ec4899);border-radius:99px;"></div>
                                </div>
                            </div>"""
                        h_px = len(items) * 54 + 56
                        components.html(f"""<!DOCTYPE html><html>
                        <head><link href="https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=DM+Sans:wght@400;600;700&display=swap" rel="stylesheet"></head>
                        <body style="margin:0;padding:4px 0;font-family:'DM Sans',sans-serif;background:white;">
                            <div style="display:flex;align-items:center;gap:8px;margin-bottom:14px;">
                                <span style="font-size:17px;">🔴</span>
                                <span style="font-size:13px;font-weight:700;color:#111;">OOS — Lager 0, najveći potencijal</span>
                                <span style="font-size:10px;font-weight:700;color:#ec4899;background:#fdf2f8;border-radius:20px;padding:2px 8px;">top 10</span>
                            </div>
                            <div style="font-size:9px;color:#ccc;display:flex;gap:10px;margin-bottom:4px;align-items:center;">
                                <span style="width:46px;"></span>
                                <span style="flex:1;text-transform:uppercase;letter-spacing:.5px;"></span>
                                <span style="text-transform:uppercase;letter-spacing:.5px;">izgubljen profit</span>
                            </div>
                            {rows_html}
                        </body></html>""", height=h_px)

                    col_rast, col_pad = st.columns(2)
                    with col_rast:
                        _render_trend_section("Rastući trendovi", "📈", "#a855f7", rastuci_list, True)
                    with col_pad:
                        _render_trend_section("Padajući trendovi", "📉", "#ec4899", padajuci_list, False)

                    st.markdown("<div style='margin:20px 0 4px 0;'></div>", unsafe_allow_html=True)

                    if engine.has_prices and len(engine.df_oos) > 0:
                        oos_k = engine.df_oos.copy()
                        if 'Lager_danas' in oos_k.columns:
                            oos_k = oos_k[oos_k['Lager_danas'] == 0]
                        oos_top = oos_k.groupby('ID KOMITENTA').agg(
                            Izgubljeno=('Izgubljeni_profit','sum'),
                            Artikala=('id artikla','nunique')
                        ).reset_index().sort_values('Izgubljeno', ascending=False).head(10)
                        oos_items = oos_top.to_dict('records')
                        oos_max = int(oos_top['Izgubljeno'].max()) if len(oos_top) > 0 else 1
                    else:
                        oos_items = []; oos_max = 1

                    col_oos2, col_empty = st.columns(2)
                    with col_oos2:
                        _render_oos_section(oos_items, oos_max)

                if engine.has_prices:
                    with tab2:
                        period_str2 = ", ".join(engine.analitika_labels) if engine.analitika_labels else "svi meseci"
                        n_mes = len(engine.analitika_labels) if engine.analitika_labels else len(engine.mesec_labels)
                        n_obj = engine.num_komitenti

                        prof = engine.df_profit_obj.copy()
                        total_bruto = int(prof['Bruto_profit'].sum())
                        total_neto = int(prof['Neto_profit'].sum())
                        total_trosak = int(prof['Trosak_mkt'].sum())
                        total_oos_izgubljen = int(engine.df_oos['Izgubljeni_profit'].sum()) if len(engine.df_oos) > 0 else 0
                        mes_trosak = total_trosak / max(n_mes, 1)
                        mes_bruto = total_bruto / max(n_mes, 1)
                        mes_neto = total_neto / max(n_mes, 1)
                        mes_oos = total_oos_izgubljen / max(n_mes, 1)

                        st.caption(f"📅 Period analize: **{period_str2}** · {n_obj} objekata · {n_mes} meseci")

                        ka, kb, kc, kd = st.columns(4)
                        def _kard(col, label, total, mes, color, prefix=""):
                            col.markdown(f"""
                            <div style="background:white;border-radius:12px;padding:16px 18px;
                                border-left:4px solid {color};box-shadow:0 2px 8px rgba(0,0,0,0.07);height:100%;">
                                <div style="font-size:10px;color:#999;font-weight:600;letter-spacing:.5px;text-transform:uppercase;margin-bottom:6px;">{label}</div>
                                <div style="font-size:22px;font-weight:700;color:{color};">{prefix}{total:,.0f} RSD</div>
                                <div style="font-size:11px;color:#aaa;margin-top:3px;">{prefix}{mes:,.0f} RSD / mesec</div>
                            </div>""", unsafe_allow_html=True)
                        _kard(ka, f"Ukupan trosak · {n_mes} meseci", total_trosak, mes_trosak, "#a855f7")
                        _kard(kb, f"Bruto profit · {n_mes} meseci", total_bruto, mes_bruto, "#10b981")
                        _kard(kc, f"Neto profit · {n_mes} meseci", total_neto, mes_neto, "#7c3aed" if total_neto > 0 else "#ec4899")
                        _kard(kd, f"OOS izgubljen · {n_mes} meseci", total_oos_izgubljen, mes_oos, "#ec4899", prefix="-")

                        st.markdown("<div style='margin:20px 0 4px 0;'></div>", unsafe_allow_html=True)

                        a_labels_trend = engine.analitika_labels if engine.analitika_labels else engine.mesec_labels
                        a_meseci_trend = engine.analitika_meseci if (engine.analitika_meseci and len(engine.analitika_meseci) > 0) else engine.meseci_order

                        bruto_po_mes = []
                        neto_po_mes = []
                        for i, lb in enumerate(a_labels_trend):
                            col_bruto = f'Bruto_{lb}'
                            col_neto = f'Neto_{lb}'
                            bruto_val = prof[col_bruto].sum() if col_bruto in prof.columns else 0
                            neto_val = prof[col_neto].sum() if col_neto in prof.columns else 0
                            bruto_po_mes.append((lb, bruto_val))
                            neto_po_mes.append((lb, neto_val))

                        def _trend_recenica(podaci, naziv):
                            vals = [v for _, v in podaci]
                            if len(vals) < 2: return ""
                            prvi_lb, prvi_v = podaci[0]
                            posl_lb, posl_v = podaci[-1]
                            if prvi_v == 0: return ""
                            promena_pct = ((posl_v - prvi_v) / abs(prvi_v)) * 100
                            smer = "porastao" if promena_pct > 0 else "pao"
                            boja = "#10b981" if promena_pct > 0 else "#ec4899"
                            return f'<span style="color:{boja};font-weight:600;">{naziv} je {smer} za {abs(promena_pct):.0f}%</span> — od <b>{prvi_v:,.0f} RSD</b> ({prvi_lb}) do <b>{posl_v:,.0f} RSD</b> ({posl_lb}).'

                        def _bar_chart_html(podaci, max_val, color_pos, color_neg):
                            bars = ""
                            for lb, val in podaci:
                                pct = abs(val) / max_val * 100 if max_val > 0 else 0
                                pct = min(pct, 100)
                                color = color_pos if val >= 0 else color_neg
                                val_fmt = f"{val:,.0f} RSD"
                                bars += f"""
                                <div style="display:flex;align-items:center;margin-bottom:5px;gap:8px;">
                                    <div style="width:52px;font-size:11px;color:#888;text-align:right;flex-shrink:0;">{lb}</div>
                                    <div style="flex:1;background:#f5f0ff;border-radius:3px;height:18px;position:relative;">
                                        <div style="width:{pct:.1f}%;background:{color};height:100%;border-radius:3px;transition:width .3s;"></div>
                                    </div>
                                    <div style="width:110px;font-size:11px;color:#555;font-weight:600;flex-shrink:0;">{val_fmt}</div>
                                </div>"""
                            return f'<div style="padding:4px 0;">{bars}</div>'

                        max_bruto = max(abs(v) for _, v in bruto_po_mes) if bruto_po_mes else 1
                        max_neto = max(abs(v) for _, v in neto_po_mes) if neto_po_mes else 1

                        col_bruto, col_neto = st.columns(2)

                        with col_bruto:
                            st.markdown('<div class="section-title">📈 Mesečni trend bruto profita</div>', unsafe_allow_html=True)
                            rec_b = _trend_recenica(bruto_po_mes, "Bruto profit")
                            if rec_b: st.markdown(f'<p style="font-size:13px;color:#555;margin-bottom:6px;">{rec_b}</p>', unsafe_allow_html=True)
                            chart_b = _bar_chart_html(bruto_po_mes, max_bruto, "#a855f7", "#ec4899")
                            components.html(f'<!DOCTYPE html><html><body style="margin:0;padding:8px 12px;font-family:sans-serif;">{chart_b}</body></html>', height=len(bruto_po_mes)*28+20)

                        with col_neto:
                            st.markdown('<div class="section-title">📉 Mesečni trend neto profita</div>', unsafe_allow_html=True)
                            rec_n = _trend_recenica(neto_po_mes, "Neto profit")
                            if rec_n: st.markdown(f'<p style="font-size:13px;color:#555;margin-bottom:6px;">{rec_n}</p>', unsafe_allow_html=True)
                            chart_n = _bar_chart_html(neto_po_mes, max_neto, "#7c3aed", "#ec4899")
                            components.html(f'<!DOCTYPE html><html><body style="margin:0;padding:8px 12px;font-family:sans-serif;">{chart_n}</body></html>', height=len(neto_po_mes)*28+20)

                        st.markdown("<div style='margin:20px 0 4px 0;'></div>", unsafe_allow_html=True)

                        st.markdown('<div class="section-title">🏪 Profitabilnost po objektima</div>', unsafe_allow_html=True)

                        ukupno_obj = len(prof)
                        neto_neg = prof[prof['Neto_profit'] <= 0]
                        n_neto_neg = len(neto_neg)
                        oos_neg = prof[(prof['Neto_profit'] <= 0) & (prof['Potencijalni_profit'] > 0)]
                        n_oos_neg = len(oos_neg)
                        pravi_neg = prof[(prof['Neto_profit'] <= 0) & (prof['Potencijalni_profit'] <= 0)]
                        n_pravi_neg = len(pravi_neg)
                        pct_pravi = round(n_pravi_neg / max(ukupno_obj, 1) * 100)
                        trosak_po_obj = engine.trosak_po_objektu
                        trosak_mes_obj = trosak_po_obj / max(n_mes, 1)
                        usteda_trosak = n_pravi_neg * trosak_po_obj
                        usteda_gubitak = abs(pravi_neg['Neto_profit'].sum()) if n_pravi_neg > 0 else 0
                        usteda_mes = (usteda_trosak + usteda_gubitak) / max(n_mes, 1)

                        n_profitabilni = ukupno_obj - n_neto_neg
                        pct_prof = n_profitabilni / max(ukupno_obj, 1)
                        pct_oos_neg_v = n_oos_neg / max(ukupno_obj, 1)
                        pct_pravi_v = n_pravi_neg / max(ukupno_obj, 1)

                        cx, cy, r_out, r_in = 110, 110, 90, 60
                        def _arc_path(cx, cy, r, start_deg, end_deg):
                            s = math.radians(start_deg - 90)
                            e = math.radians(end_deg - 90)
                            large = 1 if (end_deg - start_deg) > 180 else 0
                            x1,y1 = cx+r*math.cos(s), cy+r*math.sin(s)
                            x2,y2 = cx+r*math.cos(e), cy+r*math.sin(e)
                            return f"M {x1:.1f} {y1:.1f} A {r} {r} 0 {large} 1 {x2:.1f} {y2:.1f}"
                        def _donut_seg(cx, cy, ro, ri, start_deg, end_deg, color):
                            if end_deg - start_deg < 0.5: return ""
                            oa = _arc_path(cx, cy, ro, start_deg, end_deg)
                            s2 = math.radians(end_deg - 90); s1 = math.radians(start_deg - 90)
                            x_ie, y_ie = cx+ri*math.cos(s2), cy+ri*math.sin(s2)
                            x_is, y_is = cx+ri*math.cos(s1), cy+ri*math.sin(s1)
                            large = 1 if (end_deg - start_deg) > 180 else 0
                            x2o,y2o = cx+ro*math.cos(s2), cy+ro*math.sin(s2)
                            x1o,y1o = cx+ro*math.cos(s1), cy+ro*math.sin(s1)
                            return f'<path d="{oa} L {x_ie:.1f} {y_ie:.1f} A {ri} {ri} 0 {large} 0 {x_is:.1f} {y_is:.1f} Z" fill="{color}"/>'

                        deg_prof = pct_prof * 360
                        deg_oos = pct_oos_neg_v * 360
                        deg_pravi = pct_pravi_v * 360
                        seg1 = _donut_seg(cx, cy, r_out, r_in, 0, deg_prof, "#10b981")
                        seg2 = _donut_seg(cx, cy, r_out, r_in, deg_prof, deg_prof+deg_pravi, "#ec4899")
                        seg3 = _donut_seg(cx, cy, r_out, r_in, deg_prof+deg_pravi, deg_prof+deg_pravi+deg_oos, "#a855f7")

                        donut_svg = f"""<svg width="220" height="220" xmlns="http://www.w3.org/2000/svg">
                            {seg1}{seg2}{seg3}
                            <circle cx="{cx}" cy="{cy}" r="{r_in}" fill="white"/>
                            <text x="{cx}" y="{cy-8}" text-anchor="middle" font-size="26" font-weight="700" fill="#111" font-family="sans-serif">{n_profitabilni}</text>
                            <text x="{cx}" y="{cy+14}" text-anchor="middle" font-size="12" fill="#888" font-family="sans-serif">profitabilnih</text>
                        </svg>
                        <div style="margin-top:8px;font-size:12px;font-family:sans-serif;">
                            <div style="display:flex;align-items:center;gap:6px;margin-bottom:5px;">
                                <span style="width:12px;height:12px;background:#10b981;border-radius:2px;display:inline-block;flex-shrink:0;"></span>
                                <span style="color:#555;"><strong>{n_profitabilni} profitabilnih</strong> ({round(pct_prof*100)}% mreže)</span>
                            </div>
                            <div style="display:flex;align-items:center;gap:6px;margin-bottom:5px;">
                                <span style="width:12px;height:12px;background:#ec4899;border-radius:2px;display:inline-block;flex-shrink:0;"></span>
                                <span style="color:#555;"><strong>{n_pravi_neg} neprofitabilnih</strong> ({round(pct_pravi_v*100)}% mreže)</span>
                            </div>
                            <div style="display:flex;align-items:center;gap:6px;">
                                <span style="width:12px;height:12px;background:#a855f7;border-radius:2px;display:inline-block;flex-shrink:0;"></span>
                                <span style="color:#555;"><strong>{n_oos_neg} neto-neg. OOS</strong> potencijal</span>
                            </div>
                        </div>"""

                        tekst = f"""
    <div style="background:white;border-radius:12px;padding:20px 24px;box-shadow:0 2px 8px rgba(0,0,0,0.06);margin-bottom:16px;font-size:14px;line-height:1.8;color:#333;">
    <p>Od <strong>{ukupno_obj} objekata</strong>, <strong>{n_neto_neg}</strong> je neto negativno.
    Medjutim, <strong>{n_oos_neg}</strong> od njih ima negativan neto isključivo zbog OOS-a — kada se uračuna izgubljena zarada,
    njihov potencijal je pozitivan. Ovi objekti nisu problem, samo nisu imali robu.</p>

    <p>Pravih neprofitabilnih je <strong>{n_pravi_neg}</strong> ({pct_pravi}% ukupne mreže) — negativni čak i po potencijalu.
    Trošak po objektu je <strong>{trosak_po_obj:,.0f} RSD</strong> za {n_mes} {'mesec' if n_mes==1 else 'meseci'} /
    <strong>{trosak_mes_obj:,.0f} RSD</strong> mesečno.</p>

    <p>Zatvaranjem <strong>{n_pravi_neg} pravih neprofitabilnih</strong> skidamo trošak
    <strong>{n_pravi_neg} × {trosak_po_obj:,.0f} RSD = {usteda_trosak:,.0f} RSD</strong>
    ({usteda_trosak/max(n_mes,1):,.0f} RSD/mes) i prestajemo da gubimo
    <strong>{usteda_gubitak:,.0f} RSD</strong> ({usteda_gubitak/max(n_mes,1):,.0f} RSD/mes) na negativnim objektima.
    Ostaju samo objekti koji su u plusu.</p>
    </div>"""
                        col_tekst, col_donut = st.columns([3, 1])
                        with col_tekst:
                            st.markdown(tekst, unsafe_allow_html=True)
                        with col_donut:
                            components.html(f"""<!DOCTYPE html><html><body style="margin:0;padding:12px 8px;font-family:sans-serif;background:transparent;">
                                {donut_svg}
                            </body></html>""", height=310)

                        a_labels_trend2 = engine.analitika_labels if engine.analitika_labels else engine.mesec_labels
                        a_meseci_trend2 = engine.analitika_meseci if (engine.analitika_meseci and len(engine.analitika_meseci) > 0) else engine.meseci_order

                        chart_mes_data = []
                        for i, (lb, (g, m)) in enumerate(zip(a_labels_trend2, a_meseci_trend2)):
                            col_neto_lb = f'Neto_{lb}'
                            if col_neto_lb in prof.columns:
                                n_prof_mes = int((prof[col_neto_lb] > 0).sum())
                                n_nepr_mes = int((prof[col_neto_lb] <= 0).sum())
                            else:
                                n_prof_mes = 0; n_nepr_mes = 0
                            chart_mes_data.append((lb, n_prof_mes, n_nepr_mes))

                        if chart_mes_data:
                            max_obj_mes = max(a + b for _, a, b in chart_mes_data) if chart_mes_data else 1
                            bar_w = max(30, min(60, 700 // max(len(chart_mes_data), 1)))
                            bars_html = ""
                            for lb, np_v, nn_v in chart_mes_data:
                                h_p = int(np_v / max(max_obj_mes, 1) * 140)
                                h_n = int(nn_v / max(max_obj_mes, 1) * 140)
                                bars_html += f"""
                                <div style="display:flex;flex-direction:column;align-items:center;gap:2px;">
                                    <div style="display:flex;align-items:flex-end;gap:3px;height:160px;">
                                        <div style="width:{bar_w}px;height:{h_p}px;background:#a855f7;border-radius:3px 3px 0 0;position:relative;">
                                            <span style="position:absolute;top:-18px;left:50%;transform:translateX(-50%);font-size:10px;font-weight:700;color:#7c3aed;white-space:nowrap;">{np_v}</span>
                                        </div>
                                        <div style="width:{bar_w}px;height:{h_n}px;background:#ec4899;border-radius:3px 3px 0 0;position:relative;">
                                            <span style="position:absolute;top:-18px;left:50%;transform:translateX(-50%);font-size:10px;font-weight:700;color:#be185d;white-space:nowrap;">{nn_v}</span>
                                        </div>
                                    </div>
                                    <div style="font-size:10px;color:#888;margin-top:4px;text-align:center;width:{bar_w*2+3}px;">{lb}</div>
                                </div>"""
                            chart_html = f"""<!DOCTYPE html><html><body style="margin:0;padding:0;font-family:sans-serif;background:white;">
                            <div style="padding:16px 20px;">
                                <div style="display:flex;gap:16px;margin-bottom:14px;">
                                    <span style="display:flex;align-items:center;gap:5px;font-size:12px;color:#555;">
                                        <span style="width:12px;height:12px;background:#a855f7;border-radius:2px;display:inline-block;"></span> Profitabilni taj mesec (neto &gt; 0)
                                    </span>
                                    <span style="display:flex;align-items:center;gap:5px;font-size:12px;color:#555;">
                                        <span style="width:12px;height:12px;background:#ec4899;border-radius:2px;display:inline-block;"></span> Neprofitabilni taj mesec (neto ≤ 0)
                                    </span>
                                </div>
                                <div style="display:flex;gap:6px;align-items:flex-end;overflow-x:auto;padding-bottom:4px;">
                                    {bars_html}
                                </div>
                            </div>
                            </body></html>"""
                            components.html(chart_html, height=220)
                            st.markdown('''<p style="font-size:12px;color:#9ca3af;margin-top:4px;">
                            ℹ️ Grafikon prikazuje profitabilnost po potencijalu <strong>za svaki mesec posebno</strong> — razlikuje se od ukupnih brojeva iznad koji se odnose na <strong>ceo analizirani period</strong>. Na primer, objekat koji je u poslednjem mesecu neprofitabilan može biti profitabilan gledano kroz ceo period.
                            </p>''', unsafe_allow_html=True)

                        st.markdown("<div style='margin:20px 0 4px 0;'></div>", unsafe_allow_html=True)

                        st.markdown('<div class="section-title">🔴 OOS — Izgubljena zarada zbog nedostatka robe</div>', unsafe_allow_html=True)
                        if len(engine.df_oos) > 0:
                            a_labels_oos = engine.analitika_labels if engine.analitika_labels else engine.mesec_labels

                            oos_ukupno = int(engine.df_oos['Izgubljeni_profit'].sum())
                            oos_mes_avg = oos_ukupno // max(n_mes, 1)
                            oos_kombinacija = int((engine.df_oos['OOS_meseci'] > 0).sum()) if 'OOS_meseci' in engine.df_oos.columns else len(engine.df_oos)
                            oos_0_danas = int((engine.df_oos.get('Lager_danas', 0) == 0).sum()) if 'Lager_danas' in engine.df_oos.columns else oos_kombinacija

                            o1, o2, o3 = st.columns(3)
                            def _oos_kard(col, label, val, suffix=""):
                                col.markdown(f"""<div style="background:white;border-radius:12px;padding:16px 18px;
                                    border-top:3px solid #ec4899;box-shadow:0 2px 8px rgba(0,0,0,0.07);text-align:center;">
                                    <div style="font-size:22px;font-weight:700;color:#ec4899;">{val:,}{suffix}</div>
                                    <div style="font-size:11px;color:#aaa;margin-top:4px;text-transform:uppercase;letter-spacing:.5px;">{label}</div>
                                </div>""", unsafe_allow_html=True)
                            _oos_kard(o1, f"Izgubljen profit · {n_mes} meseci (RSD)", oos_ukupno)
                            _oos_kard(o2, "Prosečno mesečno (RSD)", oos_mes_avg)
                            _oos_kard(o3, "Kombinacija na 0 lagera danas", oos_0_danas)

                            st.markdown("<div style='margin:18px 0 4px 0;'></div>", unsafe_allow_html=True)

                            mes_izgub = []
                            mes_oos_count = []
                            for lb in a_labels_oos:
                                col_izgub = f'Izgub_{lb}'
                                col_oos = f'OOS_{lb}'
                                v_izgub = int(engine.df_oos[col_izgub].sum()) if col_izgub in engine.df_oos.columns else 0
                                v_oos = int((engine.df_oos[col_oos] > 0).sum()) if col_oos in engine.df_oos.columns else 0
                                mes_izgub.append(v_izgub)
                                mes_oos_count.append(v_oos)

                            if any(v > 0 for v in mes_izgub):
                                max_izgub = max(mes_izgub) if mes_izgub else 1
                                chart_w = 860
                                chart_h = 220
                                pad_l, pad_r, pad_t, pad_b = 60, 20, 30, 40
                                plot_w = chart_w - pad_l - pad_r
                                plot_h = chart_h - pad_t - pad_b
                                n_pts = len(a_labels_oos)

                                def px(i): return pad_l + int(i / max(n_pts-1,1) * plot_w)
                                def py(v): return pad_t + plot_h - int(v / max(max_izgub,1) * plot_h)

                                pts_area = " ".join(f"{px(i)},{py(v)}" for i, v in enumerate(mes_izgub))
                                pts_area = f"{px(0)},{pad_t+plot_h} " + pts_area + f" {px(n_pts-1)},{pad_t+plot_h}"
                                pts_line = " ".join(f"{px(i)},{py(v)}" for i, v in enumerate(mes_izgub))

                                dots = ""
                                labels_svg = ""
                                x_labels = ""
                                for i, (lb, v, vc) in enumerate(zip(a_labels_oos, mes_izgub, mes_oos_count)):
                                    x, y = px(i), py(v)
                                    v_k = f"{v//1000}k" if v >= 1000 else str(v)
                                    dots += f'<circle cx="{x}" cy="{y}" r="5" fill="#a855f7" stroke="white" stroke-width="2"/>'
                                    labels_svg += f'<text x="{x}" y="{y-10}" text-anchor="middle" font-size="10" font-weight="700" fill="#7c3aed">{v_k}</text>'
                                    labels_svg += f'<text x="{x}" y="{y+20}" text-anchor="middle" font-size="9" fill="#999">({vc})</text>'
                                    x_labels += f'<text x="{x}" y="{chart_h-6}" text-anchor="middle" font-size="9" fill="#aaa">{lb}</text>'

                                svg = f"""<svg width="{chart_w}" height="{chart_h}" xmlns="http://www.w3.org/2000/svg" style="font-family:sans-serif;">
                                    <text x="{pad_l-5}" y="{pad_t-8}" font-size="10" fill="#888">Izgubljen profit (RSD)</text>
                                    <text x="{chart_w-pad_r}" y="{pad_t-8}" font-size="10" fill="#aaa" text-anchor="end">Broj OOS kombinacija u zagradama</text>
                                    <polygon points="{pts_area}" fill="#a855f7" fill-opacity="0.08"/>
                                    <polyline points="{pts_line}" fill="none" stroke="#a855f7" stroke-width="2.5"/>
                                    {dots}{labels_svg}{x_labels}
                                </svg>"""
                                components.html(f'<!DOCTYPE html><html><body style="margin:0;padding:0;background:white;">{svg}</body></html>', height=chart_h+10)

                            oos_art = engine.df_oos.groupby(['id artikla','Naziv artikla']).agg(
                                Izgubljeni_profit=('Izgubljeni_profit','sum')
                            ).reset_index().sort_values('Izgubljeni_profit', ascending=False).head(5)

                            bar_colors = ["#a855f7","#ec4899","#7c3aed","#c084fc","#f472b6"]
                            top5_max = int(oos_art['Izgubljeni_profit'].max()) if len(oos_art) > 0 else 1
                            bars5 = ""
                            for i, (_, row) in enumerate(oos_art.iterrows()):
                                naziv = str(row['Naziv artikla'])[:35]
                                val = int(row['Izgubljeni_profit'])
                                pct = val / top5_max * 100
                                color = bar_colors[i % len(bar_colors)]
                                bars5 += f"""
                                <div style="display:flex;align-items:center;gap:10px;margin-bottom:10px;">
                                    <div style="width:200px;font-size:12px;color:#444;text-align:right;flex-shrink:0;">{naziv}</div>
                                    <div style="flex:1;background:#f5f0ff;border-radius:4px;height:22px;position:relative;">
                                        <div style="width:{pct:.1f}%;background:{color};height:100%;border-radius:4px;"></div>
                                    </div>
                                    <div style="width:110px;font-size:12px;font-weight:700;color:{color};flex-shrink:0;">{val:,} RSD</div>
                                </div>"""

                            st.markdown("**Top 5 artikala po izgubljenom profitu:**")
                            components.html(f"""<!DOCTYPE html><html><body style="margin:0;padding:8px 12px;font-family:sans-serif;background:white;">
                                {bars5}
                            </body></html>""", height=len(oos_art)*42+20)

                            with st.expander("📋 Svi artikli po izgubljenom profitu"):
                                oos_art_all = engine.df_oos.groupby(['id artikla','Naziv artikla']).agg(
                                    Objekata=('ID KOMITENTA','nunique'),
                                    OOS_meseci=('OOS_meseci','sum'),
                                    Izgubljeni_profit=('Izgubljeni_profit','sum')
                                ).reset_index().sort_values('Izgubljeni_profit', ascending=False)
                                oos_art_all.columns = ['ID Art.','Naziv','Objekata','OOS meseci','Izg. profit (RSD)']
                                st.dataframe(oos_art_all, use_container_width=True, height=300)
                        else:
                            st.success("Nema OOS problema!")

                        st.markdown("<div style='margin:24px 0 4px 0;'></div>", unsafe_allow_html=True)
                        st.markdown('<div class="section-title">⚡ Scenario: Optimalna mreža</div>', unsafe_allow_html=True)

                        prof2 = engine.df_profit_obj.copy()
                        oos_ukupno2 = int(engine.df_oos['Izgubljeni_profit'].sum()) if len(engine.df_oos) > 0 else 0

                        pozitivni = prof2[prof2['Potencijalni_profit'] > 0]
                        neto_pozitivnih = int(pozitivni['Neto_profit'].sum())

                        pravi_neg2 = prof2[(prof2['Neto_profit'] <= 0) & (prof2['Potencijalni_profit'] <= 0)]
                        n_pravi_neg2 = len(pravi_neg2)
                        usteda_trosak2 = int(n_pravi_neg2 * engine.trosak_po_objektu)
                        usteda_gubitak2 = int(abs(pravi_neg2['Neto_profit'].sum()))

                        ukupni_potencijal = neto_pozitivnih + usteda_trosak2 + usteda_gubitak2 + oos_ukupno2
                        stvarni_neto = int(prof2['Neto_profit'].sum())
                        razlika = ukupni_potencijal - stvarni_neto

                        period_sc = period_str2

                        def _red(label, val, color="#10b981", bold_val=True):
                            val_str = f"+{val:,} RSD" if val >= 0 else f"{val:,} RSD"
                            v_style = f"font-weight:{'700' if bold_val else '400'};color:{color};"
                            return f"""<div style="display:flex;justify-content:space-between;align-items:center;
                                padding:8px 0;border-bottom:1px solid #f3f4f6;">
                                <span style="font-size:13px;color:#555;">{label}</span>
                                <span style="{v_style}font-size:13px;">{val_str}</span>
                            </div>"""

                        def _red_bold(label, val, color="#111"):
                            val_str = f"= {val:,} RSD"
                            return f"""<div style="display:flex;justify-content:space-between;align-items:center;
                                padding:10px 0;border-top:2px solid #e5e7eb;margin-top:4px;">
                                <span style="font-size:14px;font-weight:700;color:#111;">{label}</span>
                                <span style="font-size:14px;font-weight:700;color:{color};">{val_str}</span>
                            </div>"""

                        scenario_html = f"""
                        <div style="background:white;border-radius:12px;padding:20px 24px;
                            box-shadow:0 2px 8px rgba(0,0,0,0.07);font-family:sans-serif;">
                            <div style="font-size:12px;font-weight:600;color:#a855f7;margin-bottom:12px;
                                text-transform:uppercase;letter-spacing:.5px;">
                                Period: {period_sc} ({n_mes} meseci)
                            </div>
                            <p style="font-size:13px;color:#666;margin-bottom:14px;">
                                Ako se istovremeno zatvore neprofitabilni objekti i eliminiše OOS, mreža ide sa
                                <strong>{stvarni_neto:,} RSD</strong> neto profita na
                                <strong style="color:#10b981;">+{ukupni_potencijal:,} RSD</strong> za {n_mes} meseci.
                            </p>
                            {_red(f"Neto profit pozitivnih objekata (potencijal > 0)", neto_pozitivnih, "#10b981")}
                            {_red(f"Ušteda: zatvaranje {n_pravi_neg2} neprofitabilnih obj.", usteda_trosak2 + usteda_gubitak2, "#10b981")}
                            {_red(f"Povraćaj izgub. zarade (OOS eliminacija)", oos_ukupno2, "#10b981")}
                            {_red_bold(f"UKUPNI POTENCIJAL ({n_mes} meseci)", ukupni_potencijal, "#10b981")}
                            <div style="height:8px;"></div>
                            {_red(f"Stvarni neto profit ({n_mes} meseci)", stvarni_neto, "#555", False)}
                            {_red(f"Razlika — potencijal koji još nije ostvaren", razlika, "#a855f7")}
                        </div>"""
                        st.markdown(scenario_html, unsafe_allow_html=True)

                        if engine.region_map:
                            st.markdown("<div style='margin:28px 0 6px 0;'></div>", unsafe_allow_html=True)
                            st.markdown('<div class="section-title">🗺️ Profitabilnost po okruzima</div>', unsafe_allow_html=True)

                            prof_reg = prof.copy()
                            prof_reg['Region'] = prof_reg['ID KOMITENTA'].map(engine.region_map).fillna('Ostalo')
                            prof_reg['Profitabilan'] = prof_reg['Neto_profit'] > 0

                            reg_grp = prof_reg.groupby('Region').agg(
                                Ukupno=('ID KOMITENTA','count'),
                                Ostaje=('Profitabilan','sum'),
                            ).reset_index()
                            reg_grp['Zatvara'] = reg_grp['Ukupno'] - reg_grp['Ostaje']
                            reg_grp = reg_grp.sort_values('Ukupno', ascending=False).reset_index(drop=True)
                            mali_okruzi_df = reg_grp[reg_grp['Ostaje'] < 5]
                            mali_okruzi = mali_okruzi_df['Region'].tolist()

                            rows_html = ""
                            for _, r in reg_grp.iterrows():
                                okrug = r['Region']
                                ukupno = int(r['Ukupno'])
                                ostaje = int(r['Ostaje'])
                                zatvara = int(r['Zatvara'])
                                mali = " *" if okrug in mali_okruzi else ""
                                mali_color = "#a855f7" if mali else "#111"
                                pct_o = ostaje / max(ukupno, 1) * 100
                                pct_z = zatvara / max(ukupno, 1) * 100
                                bar = f"""<div style="display:flex;width:120px;height:14px;border-radius:3px;overflow:hidden;">
                                    <div style="width:{pct_o:.0f}%;background:#a855f7;"></div>
                                    <div style="width:{pct_z:.0f}%;background:#ec4899;"></div>
                                </div>"""
                                rows_html += f"""<tr style="border-bottom:1px solid #f3f4f6;">
                                    <td style="padding:7px 10px;font-size:13px;color:{mali_color};font-weight:600;">{okrug}{mali}</td>
                                    <td style="padding:7px 10px;font-size:13px;font-weight:700;text-align:center;">{ukupno}</td>
                                    <td style="padding:7px 10px;font-size:13px;text-align:center;">
                                        <span style="color:#a855f7;font-weight:700;">{ostaje}</span>
                                        <span style="color:#999;"> / </span>
                                        <span style="color:#ec4899;font-weight:700;">{zatvara}</span>
                                    </td>
                                    <td style="padding:7px 16px;">{bar}</td>
                                </tr>"""

                            uk_ukupno = int(reg_grp['Ukupno'].sum())
                            uk_ostaje = int(reg_grp['Ostaje'].sum())
                            uk_zatvara = int(reg_grp['Zatvara'].sum())
                            rows_html += f"""<tr style="border-top:2px solid #e5e7eb;background:#f9fafb;">
                                <td style="padding:9px 10px;font-size:13px;font-weight:700;">UKUPNO</td>
                                <td style="padding:9px 10px;font-size:13px;font-weight:700;text-align:center;">{uk_ukupno}</td>
                                <td style="padding:9px 10px;font-size:13px;text-align:center;">
                                    <span style="color:#a855f7;font-weight:700;">{uk_ostaje}</span>
                                    <span style="color:#999;"> / </span>
                                    <span style="color:#ec4899;font-weight:700;">{uk_zatvara}</span>
                                </td>
                                <td></td>
                            </tr>"""

                            header_html = """<tr style="background:#f9fafb;border-bottom:2px solid #e5e7eb;">
                                <th style="padding:9px 10px;font-size:11px;color:#888;font-weight:600;text-align:left;text-transform:uppercase;letter-spacing:.4px;">Okrug</th>
                                <th style="padding:9px 10px;font-size:11px;color:#888;font-weight:600;text-align:center;text-transform:uppercase;letter-spacing:.4px;">Ukupno obj.</th>
                                <th style="padding:9px 10px;font-size:11px;color:#888;font-weight:600;text-align:center;text-transform:uppercase;letter-spacing:.4px;">✓ Ostaje / ✗ Zatvara</th>
                                <th style="padding:9px 10px;font-size:11px;color:#888;font-weight:600;text-transform:uppercase;letter-spacing:.4px;"></th>
                            </tr>"""

                            tbl_height = len(reg_grp) * 34 + 80
                            components.html(f"""<!DOCTYPE html><html><body style="margin:0;padding:0;font-family:sans-serif;background:white;">
                            <table style="width:100%;border-collapse:collapse;">
                                <thead>{header_html}</thead>
                                <tbody>{rows_html}</tbody>
                            </table>
                            </body></html>""", height=tbl_height)

                            if mali_okruzi:
                                mali_str = ", ".join(mali_okruzi)
                                st.markdown(f'<div style="font-size:12px;color:#a855f7;padding:6px 4px;">* Okruzi sa manje od 5 profitabilnih objekata ({mali_str}): Ne preporučuje se angazovanje komercijalistu isključivo za ove okruge — broj preostalih objekata premali je da bi opravdao redovne obilaske.</div>', unsafe_allow_html=True)

                            if len(mali_okruzi_df) > 0:
                                st.markdown("<div style='margin:20px 0 6px 0;'></div>", unsafe_allow_html=True)

                                prof_reg_mali = prof_reg[prof_reg['Region'].isin(mali_okruzi) & (prof_reg['Neto_profit'] > 0)]
                                n_mali_prof = len(prof_reg_mali)
                                neto_mali_prof = int(prof_reg_mali['Neto_profit'].sum())
                                usteda_mali_trosak = int(n_mali_prof * engine.trosak_po_objektu)

                                scA_potencijal = ukupni_potencijal
                                scB_potencijal = scA_potencijal + usteda_mali_trosak - neto_mali_prof

                                period_label = period_str2

                                def _sc_red(label, val, color="#555", bold=False):
                                    sign = "+" if val >= 0 else ""
                                    fw = "700" if bold else "400"
                                    return f"""<div style="display:flex;justify-content:space-between;padding:7px 0;border-bottom:1px solid #f3f4f6;">
                                        <span style="font-size:13px;color:#555;">{label}</span>
                                        <span style="font-size:13px;font-weight:{fw};color:{color};">{sign}{val:,} RSD</span>
                                    </div>"""

                                def _sc_total(label, val, color="#10b981"):
                                    return f"""<div style="display:flex;justify-content:space-between;padding:9px 0;border-top:2px solid #e5e7eb;margin-top:4px;">
                                        <span style="font-size:14px;font-weight:700;color:#111;">{label}</span>
                                        <span style="font-size:14px;font-weight:700;color:{color};">= {val:,} RSD</span>
                                    </div>"""

                                sc_html = f"""<div style="font-family:sans-serif;background:white;border-radius:12px;
                                    padding:20px 24px;box-shadow:0 2px 8px rgba(0,0,0,0.07);">
                                    <div style="font-size:12px;font-weight:600;color:#a855f7;text-transform:uppercase;
                                        letter-spacing:.5px;margin-bottom:14px;">
                                        Uticaj zatvaranja objekata u malim okruzima ({period_label})
                                    </div>
                                    <p style="font-size:13px;color:#666;margin-bottom:14px;">
                                        Zatvaranjem {n_mali_prof} profitabilnih objekata u {len(mali_okruzi)} malih okruga
                                        štedimo trošak, ali gubimo deo zarade. Poređenje dva scenarija:
                                    </p>

                                    <div style="font-size:12px;font-weight:600;color:#7c3aed;margin:10px 0 6px 0;">
                                        Scenario A: Zatvaramo samo {n_pravi_neg2} neprofitabilnih + OOS eliminacija
                                    </div>
                                    {_sc_red(f"Neto profit pozitivnih objekata ({n_mes}m)", neto_pozitivnih, "#10b981", False)}
                                    {_sc_red(f"Ušteda: zatvaranje {n_pravi_neg2} neprofitabilnih ({n_mes}m)", usteda_trosak2 + usteda_gubitak2, "#10b981", False)}
                                    {_sc_red(f"Povraćaj OOS izgubljene zarade ({n_mes}m)", oos_ukupno2, "#10b981", False)}
                                    {_sc_total(f"POTENCIJAL SCENARIO A", scA_potencijal)}

                                    <div style="font-size:12px;font-weight:600;color:#ec4899;margin:16px 0 6px 0;">
                                        Scenario B: Scenario A + zatvaramo i {n_mali_prof} obj. iz malih okruga
                                    </div>
                                    {_sc_red(f"Potencijal Scenario A", scA_potencijal, "#10b981", False)}
                                    {_sc_red(f"Ušteda troška: {n_mali_prof} obj. × {engine.trosak_po_objektu:,.0f} RSD × {n_mes} mes", usteda_mali_trosak, "#10b981", False)}
                                    {_sc_red(f"Izgubljen profit zatvorenih {n_mali_prof} obj. ({n_mes}m)", -neto_mali_prof, "#ec4899", False)}
                                    {_sc_total(f"POTENCIJAL SCENARIO B", scB_potencijal, "#10b981" if scB_potencijal >= scA_potencijal else "#a855f7")}
                                </div>"""
                                components.html(f'<!DOCTYPE html><html><body style="margin:0;padding:0;">{sc_html}</body></html>', height=420)

                st.markdown("---")
                excel_buf = create_excel(engine)
                fname_xl = f"ANALITIKA_{datetime.date.today().strftime('%Y%m%d')}.xlsx"
                st.download_button(f"📥 Preuzmi Excel — {fname_xl}", data=excel_buf, file_name=fname_xl,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

            except Exception as e:
                st.error(f"Greska: {str(e)}")
                import traceback; st.code(traceback.format_exc())

    else:
        # --- POČETNA STRANICA BEZ FAJLA ---
        components.html("""<!DOCTYPE html><html><head>
<link href="https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;600;700&display=swap" rel="stylesheet">
</head><body style="margin:0;padding:0;background:transparent;font-family:'Poppins',sans-serif;">
<div style="max-width:680px;margin:32px auto 0 auto;padding:0 16px;">
  <p style="font-size:11px;color:#9ca3af;font-weight:600;letter-spacing:1.5px;text-transform:uppercase;margin-bottom:14px;">
    AMAN d.o.o. &middot; Analiticki sistem
  </p>
  <h1 style="font-size:36px;font-weight:700;color:#1a0533;line-height:1.2;margin-bottom:12px;margin-top:0;">
    Predikcija prodaje<br>
    <span style="background:linear-gradient(135deg,#a855f7,#ec4899);-webkit-background-clip:text;-webkit-text-fill-color:transparent;">
      &amp; Porudzbine
    </span>
  </h1>
  <p style="font-size:15px;color:#6b7280;margin-bottom:28px;line-height:1.6;">
    Profitabilnost objekata &middot; OOS analiza &middot; Trendovi komitenata &middot; Analiza akcije
  </p>
  <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:36px;">
    <span style="font-size:12px;background:rgba(168,85,247,0.10);color:#7c3aed;border-radius:99px;padding:5px 14px;font-weight:600;">Predikcija</span>
    <span style="font-size:12px;background:rgba(236,72,153,0.09);color:#be185d;border-radius:99px;padding:5px 14px;font-weight:600;">Profitabilnost</span>
    <span style="font-size:12px;background:rgba(239,68,68,0.09);color:#b91c1c;border-radius:99px;padding:5px 14px;font-weight:600;">OOS analiza</span>
    <span style="font-size:12px;background:rgba(16,185,129,0.09);color:#065f46;border-radius:99px;padding:5px 14px;font-weight:600;">Trendovi</span>
  </div>
  <div style="height:1px;background:linear-gradient(90deg,rgba(168,85,247,0.3),rgba(236,72,153,0.2),transparent);margin-bottom:28px;"></div>
  <p style="font-size:14px;color:#9ca3af;text-align:center;margin-top:8px;">
    &#8592; Ucitaj Excel fajl u levom panelu da pocnes analizu
  </p>
</div>
</body></html>""", height=340)

# ============================================================
# MESECNI IZVESTAJ PRODAJE
# ============================================================
elif page == 'mesecni':
    render_header("Mesečni izveštaj prodaje · Sistemi · Zalihe")

    @st.cache_data(ttl=300)
    def build_mesecni_html():
        import html as html_mod, json as json_mod

        buf_s = load_github_excel("tabela sistemi3.xlsx")
        buf_t = load_github_excel("TABELA TROSKOVA.xlsx")
        if buf_s is None:
            return None
        if buf_t is None:
            return None
        cfg = load_github_config()
        iskljuci_poslednji = not cfg.get("ukljuci_poslednji_mesec", False)

        df = pd.read_excel(buf_s, sheet_name='tabela')
        df.columns = df.columns.astype(str).str.strip()
        df_mag = pd.read_excel(buf_s, sheet_name='zalihe magacin ')
        df_mag.columns = df_mag.columns.astype(str).str.strip()
        for c in df_mag.columns:
            if 'kol' in c.lower().strip(): df_mag.rename(columns={c:'KOL'}, inplace=True); break

        df_troskovi = pd.read_excel(buf_t, sheet_name='tabela')
        df_troskovi.columns = df_troskovi.columns.str.strip()
        df_troskovi['SISTEM'] = df_troskovi['SISTEM'].str.strip()
        df_troskovi['Mesec'] = pd.to_numeric(df_troskovi['Mesec'], errors='coerce').fillna(0).astype(int)
        df_troskovi['Godina'] = pd.to_numeric(df_troskovi['Godina'], errors='coerce').fillna(0).astype(int)

        trosak_kolone = [c for c in df_troskovi.columns if c not in ['SISTEM','Mesec','Godina']]
        trosak_nazivi = {
            'trosak transporta':'Trošak transporta','trosak advokata':'Trošak advokata',
            'tehnomedia/gigatron':'Tehnomedia / Gigatron','trosak vozila':'Trošak vozila',
            'reprezentacija':'Reprezentacija','trosak programa':'Trošak programa',
            'racuni + banka+ osiguranje':'Računi + banka + osiguranje','plate':'Plate',
            'trosak laboratorije':'Trošak laboratorije','promotivni troskovi':'Promotivni troškovi',
            'troškovi kanc/ materijala':'Troškovi kanc. materijala','ostali troškovi':'Ostali troškovi',
        }
        mapa_troskova = {}
        for _, row in df_troskovi.iterrows():
            s=row['SISTEM']; g=int(row['Godina']); m=int(row['Mesec'])
            key=f"{s}|{g}-{m}"; mapa_troskova[key]={}
            for c in trosak_kolone:
                mapa_troskova[key][c]=float(row[c]) if pd.notna(row[c]) else 0

        col_nacin = next((c for c in df.columns if "PLACANJA" in c.upper().replace('Č','C').replace('Ć','C')), None)
        mapa_placanja = {}
        if col_nacin:
            temp=df[['SISTEM',col_nacin]].dropna()
            mapa_placanja=dict(zip(temp['SISTEM'],temp[col_nacin]))

        df['Mesec']=pd.to_numeric(df['Mesec'],errors='coerce').fillna(0).astype(int)
        df['Godina']=pd.to_numeric(df['Godina'],errors='coerce').fillna(0).astype(int)
        df_clean=df[(df['Godina']>=2025)&(df['Mesec']>=1)].copy()

        periodi=df_clean[['Godina','Mesec']].drop_duplicates().sort_values(['Godina','Mesec']).values.tolist()
        mapa_meseci_={1:'Jan',2:'Feb',3:'Mar',4:'Apr',5:'Maj',6:'Jun',7:'Jul',8:'Avg',9:'Sep',10:'Okt',11:'Nov',12:'Dec'}
        nazivi=[f"{mapa_meseci_[int(m)]} {str(int(g))[-2:]}" for g,m in periodi]

        poslednji_g,poslednji_m=periodi[-1]
        poslednji_naziv=f"{mapa_meseci_[int(poslednji_m)]} {int(poslednji_g)}"

        if iskljuci_poslednji:
            periodi_profit=periodi[:-1]; nazivi_profit=nazivi[:-1]
            badge_html=f'<span class="badge br">{poslednji_naziv}: ISKLJUČEN</span>'
        else:
            periodi_profit=periodi[:]; nazivi_profit=nazivi[:]
            badge_html=f'<span class="badge bg">{poslednji_naziv}: UKLJUČEN</span>'

        num_s=len(nazivi); num_p=len(nazivi_profit)

        def get_color(price):
            cm={1390:'#90EE90',1300:'#FFD1DC',1290:'#FFB6C1',1190:'#FF69B4',990:'#C71585',890:'#90EE90',800:'#FFD1DC',790:'#FFB6C1',730:'#FF69B4',690:'#C71585'}
            if not price or price==0: return None
            best=min(cm.keys(),key=lambda k:abs(k-price))
            return cm[best] if abs(best-price)<30 else None
        def esc(t): return html_mod.escape(str(t))
        def fmtnum(v): return f"{round(v):,}"
        def cell_c(bg,val):
            if not bg: return f'<td class="n">{fmtnum(val)}</td>'
            tc='#fff' if bg in ['#C71585','#FF69B4'] else '#1a1a2e'
            return f'<td class="n" style="background:{bg};color:{tc};font-weight:600">{fmtnum(val)}</td>'

        sistemi_lista=sorted(df_clean['SISTEM'].dropna().unique())
        sve_grupe=sorted(df_clean['Grupa artikla'].dropna().astype(str).str.strip().unique())

        prodaja_data_js={}
        for sistem in sistemi_lista:
            s_data=df_clean[df_clean['SISTEM']==sistem]; prodaja_data_js[sistem]={}
            for grupa in sorted(s_data['Grupa artikla'].dropna().astype(str).str.strip().unique()):
                g_data=s_data[s_data['Grupa artikla'].astype(str).str.strip()==grupa]; gvals=[]
                for g,m in periodi:
                    mask=(g_data['Godina']==int(g))&(g_data['Mesec']==int(m))
                    gvals.append(round(float(g_data.loc[mask,'Prodata kolicina ka krajnjem kupcu'].sum())))
                prodaja_data_js[sistem][grupa]=gvals

        profit_data_js={}; trosak_names_list=['Troškovi marketinga']; trosak_ids_list=['mkt']
        for ki,kat in enumerate(trosak_kolone):
            trosak_ids_list.append(f"t{ki}"); trosak_names_list.append(trosak_nazivi.get(kat,kat))
        for sistem in sistemi_lista:
            s_data=df_clean[df_clean['SISTEM']==sistem]
            status=mapa_placanja.get(sistem,None); is_f=str(status) in ['1','1.0']
            pn=[]; mn=[]
            for g,m in periodi_profit:
                mask=(s_data['Godina']==int(g))&(s_data['Mesec']==int(m))
                pn.append(round(float(s_data.loc[mask,'PROFIT3' if is_f else 'Profit'].sum())))
                mn.append(round(float(s_data.loc[mask,'MESECNI TROSAK1'].sum())))
            profit_data_js[sistem]={'profit':pn,'mkt':mn}
            for ki,kat in enumerate(trosak_kolone):
                kn=[]
                for g,m in periodi_profit:
                    key=f"{sistem}|{int(g)}-{int(m)}"
                    kn.append(round(mapa_troskova.get(key,{}).get(kat,0)))
                profit_data_js[sistem][f"t{ki}"]=kn

        drv_data_js={}
        for sistem in sistemi_lista:
            s_data=df_clean[df_clean['SISTEM']==sistem]; pn=[]; mn=[]
            for g,m in periodi_profit:
                mask=(s_data['Godina']==int(g))&(s_data['Mesec']==int(m))
                pn.append(round(float(s_data.loc[mask,'Profit'].sum())))
                mn.append(round(float(s_data.loc[mask,'MESECNI TROSAK1'].sum())))
            drv_data_js[sistem]={'profit':pn,'mkt':mn}

        # Magacin
        last3=periodi[-3:]
        df_art=df_clean[df_clean['Artikl'].notna()].copy()
        df_art['Artikl']=df_art['Artikl'].astype(str).str.strip()
        art_monthly_sales={}
        for art in df_art['Artikl'].unique():
            a_data=df_art[df_art['Artikl']==art]; sales=[]
            for g,m in last3:
                mask=(a_data['Godina']==int(g))&(a_data['Mesec']==int(m))
                sales.append(float(a_data.loc[mask,'Prodata kolicina ka krajnjem kupcu'].sum()))
            art_monthly_sales[art]=round(sum(sales)/len(sales)) if sales else 0

        def get_grupa(art_name):
            if 'HQD' in art_name.upper(): return 'HQD 1000'
            if '2000' in art_name: return 'NERD 2000'
            if 'E-cigareta' in art_name or '1000' in art_name: return 'NERD 1000'
            return 'Ostalo'

        mag_data=[]; total_mag=0; total_avg=0
        for _,row in df_mag.iterrows():
            art_name=str(row['Naziv artikla']).strip(); kol=int(row['KOL'])
            avg_sale=art_monthly_sales.get(art_name,0)
            daily=avg_sale/30 if avg_sale>0 else 0
            days=round(kol/daily) if daily>0 else 9999
            mag_data.append((art_name,kol,avg_sale,days))
            total_mag+=kol; total_avg+=avg_sale

        mag_data.sort(key=lambda x:(get_grupa(x[0]),x[0]))
        mag_rows=[]
        for art_name,kol,avg_sale,days in mag_data:
            grupa=get_grupa(art_name)
            if days<=30: dcls='style="background:#fee2e2;color:#b91c1c;font-weight:700"'
            elif days<=60: dcls='style="background:#fef3c7;color:#92400e;font-weight:600"'
            elif days<=90: dcls='style="background:#e0f2fe;color:#075985;font-weight:600"'
            else: dcls='style="color:var(--grn);font-weight:600"'
            days_str=f"{days}" if days<9999 else "∞"
            months_str=f"{days/30:.1f}" if days<9999 else "∞"
            r=f'<tr class="mag-row" data-grupa="{esc(grupa)}"><td style="font-size:10px;color:var(--t2)">{esc(art_name)}</td><td class="nb" style="color:var(--t2);font-size:10px">{esc(grupa)}</td><td class="nb">{fmtnum(kol)}</td><td class="nb">{fmtnum(avg_sale)}</td><td class="nb" {dcls}>{days_str}</td><td class="nb" {dcls}>{months_str}</td></tr>'
            mag_rows.append(r)
        total_daily=total_avg/30 if total_avg>0 else 0
        total_days=round(total_mag/total_daily) if total_daily>0 else 9999
        total_months_str=f"{total_days/30:.1f}"
        mag_rows.append(f'<tr class="totalrow"><td class="total-label">UKUPNO MAGACIN</td><td></td><td class="nb total-cell">{fmtnum(total_mag)}</td><td class="nb total-cell">{fmtnum(total_avg)}</td><td class="nb total-cell" style="font-weight:800">{total_days}</td><td class="nb total-cell" style="font-weight:800">{total_months_str}</td></tr>')

        mag_grupe_summary={}
        for art_name,kol,avg_sale,days in mag_data:
            g=get_grupa(art_name)
            if g not in mag_grupe_summary: mag_grupe_summary[g]={'kol':0,'avg':0}
            mag_grupe_summary[g]['kol']+=kol; mag_grupe_summary[g]['avg']+=avg_sale

        def get_last_zalihe(data_, periodi_list):
            for g,m in reversed(periodi_list):
                mask=(data_['Godina']==int(g))&(data_['Mesec']==int(m))
                subset=data_.loc[mask]; zal=subset['Zalihe'].sum()
                if pd.notna(zal) and zal>0:
                    return round(float(zal)), f"{mapa_meseci_[int(m)]} {str(int(g))[-2:]}"
            return 0, ""

        zal_rows=[]; si=0; zal_grand=0
        for sistem in sistemi_lista:
            si+=1; sid=f"zs{si}"; s_data=df_clean[df_clean['SISTEM']==sistem]
            s_zal,s_per=get_last_zalihe(s_data,periodi); zal_grand+=s_zal
            r=f'<tr class="sr" data-sistem="{esc(sistem)}" onclick="tog(\'{sid}\')" style="cursor:pointer"><td><button class="be" id="b-{sid}">+</button><span class="sn">{esc(sistem)}</span></td><td class="nb">{fmtnum(s_zal)}</td><td class="n" style="color:var(--t3);font-size:10px">{s_per}</td></tr>'
            zal_rows.append(r)
            grupe=sorted(s_data['Grupa artikla'].dropna().astype(str).str.strip().unique()); gi=0
            for grupa in grupe:
                gi+=1; gid=f"{sid}g{gi}"; g_data=s_data[s_data['Grupa artikla'].astype(str).str.strip()==grupa]
                g_zal,g_per=get_last_zalihe(g_data,periodi)
                r=f'<tr class="gr hidden" data-p="{sid}" data-sistem="{esc(sistem)}" data-grupa="{esc(grupa)}" onclick="tog(\'{gid}\');event.stopPropagation()" style="cursor:pointer"><td style="padding-left:28px"><button class="be beg" id="b-{gid}">+</button><span class="gn">{esc(grupa)}</span></td><td class="nb">{fmtnum(g_zal)}</td><td class="n" style="color:var(--t3);font-size:10px">{g_per}</td></tr>'
                zal_rows.append(r)
                for art in sorted(g_data['Artikl'].dropna().astype(str).str.strip().unique()):
                    if not art: continue
                    a_data=g_data[g_data['Artikl'].astype(str).str.strip()==art]
                    a_zal,a_per=get_last_zalihe(a_data,periodi)
                    r=f'<tr class="ar hidden" data-p="{gid}" data-sistem="{esc(sistem)}" data-grupa="{esc(grupa)}"><td class="an">{esc(art)}</td><td class="n" style="color:#7c8494;font-size:10px">{fmtnum(a_zal) if a_zal>0 else ""}</td><td class="n" style="color:var(--t3);font-size:9px">{a_per}</td></tr>'
                    zal_rows.append(r)
            zal_rows.append(f'<tr class="sep" data-sistem="{esc(sistem)}"><td colspan="999"></td></tr>')
        zal_rows.append(f'<tr class="totalrow" id="zalihe-total"><td class="total-label">TOTAL</td><td class="nb total-cell" style="font-size:13px">{fmtnum(zal_grand)}</td><td></td></tr>')

        prodaja_rows=[]; grand_totals=[0]*num_s; si=0
        for sistem in sistemi_lista:
            si+=1; sid=f"ps{si}"; s_data=df_clean[df_clean['SISTEM']==sistem]; vals=[]
            for g,m in periodi:
                mask=(s_data['Godina']==int(g))&(s_data['Mesec']==int(m))
                vals.append(round(float(s_data.loc[mask,'Prodata kolicina ka krajnjem kupcu'].sum())))
            total=sum(vals)
            for i in range(num_s): grand_totals[i]+=vals[i]
            r=f'<tr class="sr" data-sistem="{esc(sistem)}" onclick="tog(\'{sid}\')" style="cursor:pointer"><td><button class="be" id="b-{sid}">+</button><span class="sn">{esc(sistem)}</span></td>'
            for v in vals: r+=f'<td class="nb">{fmtnum(v)}</td>'
            r+=f'<td class="nt">{fmtnum(total)}</td></tr>'; prodaja_rows.append(r)
            grupe=sorted(s_data['Grupa artikla'].dropna().astype(str).str.strip().unique()); gi=0
            for grupa in grupe:
                gi+=1; gid=f"{sid}g{gi}"; g_data=s_data[s_data['Grupa artikla'].astype(str).str.strip()==grupa]; gvals=[]; gcene=[]
                for g,m in periodi:
                    mask=(g_data['Godina']==int(g))&(g_data['Mesec']==int(m))
                    gvals.append(round(float(g_data.loc[mask,'Prodata kolicina ka krajnjem kupcu'].sum())))
                    ac=g_data.loc[mask,'FINALNA MP'].mean()
                    gcene.append(round(float(ac)) if pd.notna(ac) and ac>0 else 0)
                gtotal=sum(gvals)
                r=f'<tr class="gr hidden" data-p="{sid}" data-sistem="{esc(sistem)}" data-grupa="{esc(grupa)}" onclick="tog(\'{gid}\');event.stopPropagation()" style="cursor:pointer"><td style="padding-left:28px"><button class="be beg" id="b-{gid}">+</button><span class="gn">{esc(grupa)}</span></td>'
                for i,v in enumerate(gvals): r+=cell_c(get_color(gcene[i]),v)
                r+=f'<td class="nb">{fmtnum(gtotal)}</td></tr>'; prodaja_rows.append(r)
            prodaja_rows.append(f'<tr class="sep" data-sistem="{esc(sistem)}"><td colspan="999"></td></tr>')
        gtt=sum(grand_totals)
        tr_r='<tr class="totalrow" id="prodaja-total"><td class="total-label">TOTAL</td>'
        for v in grand_totals: tr_r+=f'<td class="nb total-cell">{fmtnum(v)}</td>'
        tr_r+=f'<td class="nt total-cell" style="font-size:13px">{fmtnum(gtt)}</td></tr>'; prodaja_rows.append(tr_r)

        profit_rows=[]; si=0; profit_grand_totals=[0]*num_p
        for sistem in sistemi_lista:
            si+=1; sid2=f"pr{si}"; pd_s=profit_data_js[sistem]
            status=mapa_placanja.get(sistem,None); is_f=str(status) in ['1','1.0']; is_o=str(status) in ['0','0.0']
            profit_niz=pd_s['profit']; mkt_niz=pd_s['mkt']
            uk_niz=list(mkt_niz)
            for tid in trosak_ids_list[1:]:
                for j in range(num_p): uk_niz[j]+=pd_s.get(tid,[0]*num_p)[j]
            neto_niz=[profit_niz[j]-uk_niz[j] for j in range(num_p)]
            for j in range(num_p): profit_grand_totals[j]+=neto_niz[j]
            nacin="po fakturi" if is_f else ("po odjavi" if is_o else "—"); ncls="nf" if is_f else "no"
            r=f'<tr class="sr profit-sr" data-sistem="{esc(sistem)}" data-sid="{sid2}" onclick="tog(\'{sid2}\')" style="cursor:pointer"><td><button class="be" id="b-{sid2}">+</button><span class="sn">{esc(sistem)}</span></td><td class="{ncls}">{nacin}</td><td class="nb">Neto profit</td>'
            for v in neto_niz:
                cls="np" if v>=0 else "nn"; r+=f'<td class="nb {cls}">{fmtnum(v)}</td>'
            neto_total=sum(neto_niz); ncl="np" if neto_total>=0 else "nn"
            r+=f'<td class="nb {ncl}" style="font-size:13px">{fmtnum(neto_total)}</td></tr>'; profit_rows.append(r)
            r=f'<tr class="cr hidden" data-p="{sid2}" data-sistem="{esc(sistem)}"><td></td><td></td><td class="cl" style="color:var(--t2);font-style:normal;font-weight:600">{"Profit (promet)" if is_f else "Profit (odjava)"}</td>'
            for v in profit_niz: r+=f'<td class="n" style="font-weight:600;color:var(--t2)">{fmtnum(v)}</td>'
            r+=f'<td class="n" style="font-weight:700;color:var(--t2)">{fmtnum(sum(profit_niz))}</td></tr>'; profit_rows.append(r)
            r=f'<tr class="cr hidden" data-p="{sid2}" data-sistem="{esc(sistem)}" data-trosak="mkt"><td></td><td></td><td class="cl">Troškovi marketinga</td>'
            for v in mkt_niz: r+=f'<td class="n cc">{fmtnum(v)}</td>'
            r+=f'<td class="n cc" style="font-weight:700">{fmtnum(sum(mkt_niz))}</td></tr>'; profit_rows.append(r)
            for ki,kat in enumerate(trosak_kolone):
                tid=f"t{ki}"; vals=pd_s.get(tid,[0]*num_p); naz=trosak_nazivi.get(kat,kat)
                r=f'<tr class="cr hidden" data-p="{sid2}" data-sistem="{esc(sistem)}" data-trosak="{tid}"><td></td><td></td><td class="cl">{esc(naz)}</td>'
                for v in vals: r+=f'<td class="n cc">{fmtnum(v)}</td>'
                r+=f'<td class="n cc" style="font-weight:700">{fmtnum(sum(vals))}</td></tr>'; profit_rows.append(r)
            r=f'<tr class="ctr hidden ukupni-row" data-p="{sid2}" data-sistem="{esc(sistem)}" data-sid="{sid2}"><td></td><td></td><td class="ctl">UKUPNI TROŠKOVI</td>'
            for v in uk_niz: r+=f'<td class="n ctc">{fmtnum(v)}</td>'
            r+=f'<td class="n ctc" style="font-size:12px">{fmtnum(sum(uk_niz))}</td></tr>'; profit_rows.append(r)
            r=f'<tr class="nr hidden neto-row" data-p="{sid2}" data-sistem="{esc(sistem)}" data-sid="{sid2}"><td></td><td></td><td class="nl">NETO PROFIT</td>'
            for v in neto_niz:
                cls="np" if v>=0 else "nn"; r+=f'<td class="n {cls}">{fmtnum(v)}</td>'
            r+=f'<td class="n {ncl}" style="font-size:13px">{fmtnum(neto_total)}</td></tr>'; profit_rows.append(r)
            profit_rows.append(f'<tr class="sep" data-sistem="{esc(sistem)}"><td colspan="999"></td></tr>')
        pgt=sum(profit_grand_totals); pgcl="np" if pgt>=0 else "nn"
        tr_p='<tr class="totalrow" id="profit-total"><td class="total-label">TOTAL</td><td></td><td></td>'
        for v in profit_grand_totals:
            cls="np" if v>=0 else "nn"; tr_p+=f'<td class="nb total-cell {cls}">{fmtnum(v)}</td>'
        tr_p+=f'<td class="nb total-cell {pgcl}" style="font-size:13px">{fmtnum(pgt)}</td></tr>'; profit_rows.append(tr_p)

        drv_rows=[]; si=0; drv_grand_totals=[0]*num_p
        for sistem in sistemi_lista:
            si+=1; sid3=f"dv{si}"; dv=drv_data_js[sistem]
            profit_niz=dv['profit']; mkt_niz=dv['mkt']
            neto_niz=[profit_niz[j]-mkt_niz[j] for j in range(num_p)]
            for j in range(num_p): drv_grand_totals[j]+=neto_niz[j]
            r=f'<tr class="sr" data-sistem="{esc(sistem)}" data-sid="{sid3}" onclick="tog(\'{sid3}\')" style="cursor:pointer"><td><button class="be" id="b-{sid3}">+</button><span class="sn">{esc(sistem)}</span></td><td class="no">po odjavi</td><td class="nb">Neto profit</td>'
            for v in neto_niz:
                cls="np" if v>=0 else "nn"; r+=f'<td class="nb {cls}">{fmtnum(v)}</td>'
            neto_total=sum(neto_niz); ncl="np" if neto_total>=0 else "nn"
            r+=f'<td class="nb {ncl}" style="font-size:13px">{fmtnum(neto_total)}</td></tr>'; drv_rows.append(r)
            r=f'<tr class="cr hidden" data-p="{sid3}" data-sistem="{esc(sistem)}"><td></td><td></td><td class="cl" style="color:var(--t2);font-style:normal;font-weight:600">Profit (odjava)</td>'
            for v in profit_niz: r+=f'<td class="n" style="font-weight:600;color:var(--t2)">{fmtnum(v)}</td>'
            r+=f'<td class="n" style="font-weight:700;color:var(--t2)">{fmtnum(sum(profit_niz))}</td></tr>'; drv_rows.append(r)
            r=f'<tr class="cr hidden" data-p="{sid3}" data-sistem="{esc(sistem)}" data-trosak="mkt"><td></td><td></td><td class="cl">Troškovi marketinga</td>'
            for v in mkt_niz: r+=f'<td class="n cc">{fmtnum(v)}</td>'
            r+=f'<td class="n cc" style="font-weight:700">{fmtnum(sum(mkt_niz))}</td></tr>'; drv_rows.append(r)
            r=f'<tr class="ctr hidden" data-p="{sid3}" data-sistem="{esc(sistem)}" data-sid="{sid3}"><td></td><td></td><td class="ctl">UKUPNI TROŠKOVI</td>'
            for v in mkt_niz: r+=f'<td class="n ctc">{fmtnum(v)}</td>'
            r+=f'<td class="n ctc" style="font-size:12px">{fmtnum(sum(mkt_niz))}</td></tr>'; drv_rows.append(r)
            r=f'<tr class="nr hidden" data-p="{sid3}" data-sistem="{esc(sistem)}" data-sid="{sid3}"><td></td><td></td><td class="nl">NETO PROFIT</td>'
            for v in neto_niz:
                cls="np" if v>=0 else "nn"; r+=f'<td class="n {cls}">{fmtnum(v)}</td>'
            r+=f'<td class="n {ncl}" style="font-size:13px">{fmtnum(neto_total)}</td></tr>'; drv_rows.append(r)
            drv_rows.append(f'<tr class="sep" data-sistem="{esc(sistem)}"><td colspan="999"></td></tr>')
        dgt=sum(drv_grand_totals); dgcl="np" if dgt>=0 else "nn"
        tr_d='<tr class="totalrow" id="drv-total"><td class="total-label">TOTAL</td><td></td><td></td>'
        for v in drv_grand_totals:
            cls="np" if v>=0 else "nn"; tr_d+=f'<td class="nb total-cell {cls}">{fmtnum(v)}</td>'
        tr_d+=f'<td class="nb total-cell {dgcl}" style="font-size:13px">{fmtnum(dgt)}</td></tr>'; drv_rows.append(tr_d)

        ph=f'<tr><th style="text-align:left;min-width:280px">SISTEM / GRUPA</th>'
        for n in nazivi: ph+=f'<th>{n}</th>'
        ph+='<th>TOTAL</th></tr>'
        prh=f'<tr><th style="text-align:left;min-width:200px">SISTEM</th><th>NAČIN</th><th>STAVKA</th>'
        for n in nazivi_profit: prh+=f'<th>{n}</th>'
        prh+='<th>TOTAL</th></tr>'
        mh='<tr><th style="text-align:left;min-width:300px">ARTIKAL</th><th>GRUPA</th><th>MAGACIN (kom)</th><th>Ø PRODAJA/mes</th><th>DANA ZALIHA</th><th>MESECI</th></tr>'
        zsh='<tr><th style="text-align:left;min-width:280px">SISTEM / GRUPA / ARTIKAL</th><th>ZALIHE (kom)</th><th>PERIOD</th></tr>'

        info=f"{len(sistemi_lista)} sistema · {len(nazivi)} meseci"
        sistem_options=''.join([f'<option value="{esc(s)}">{esc(s)}</option>' for s in sistemi_lista])
        grupa_options=''.join([f'<option value="{esc(g)}">{esc(g)}</option>' for g in sve_grupe])
        trosak_checks=''
        for tid,tname in zip(trosak_ids_list,trosak_names_list):
            trosak_checks+=f'<label class="tcb"><input type="checkbox" checked value="{tid}" onchange="applyProfitFilters()"><span>{tname}</span></label>\n    '

        prodaja_json=json_mod.dumps(prodaja_data_js,ensure_ascii=False)
        profit_json=json_mod.dumps(profit_data_js,ensure_ascii=False)
        drv_json=json_mod.dumps(drv_data_js,ensure_ascii=False)

        last3_names=[nazivi[i] for i in range(-3,0)]
        mag_info=f"Prosek prodaje: {', '.join(last3_names)}"
        cards_html=f'''<div class="mag-cards">
          <div class="mag-card"><div class="mc-label">Ukupno magacin</div><div class="mc-val" style="color:var(--ac)">{fmtnum(total_mag)}</div><div class="mc-sub">komada</div></div>
          <div class="mag-card"><div class="mc-label">Ø Mesečna prodaja</div><div class="mc-val" style="color:var(--t1)">{fmtnum(total_avg)}</div><div class="mc-sub">kom/mesec</div></div>
          <div class="mag-card"><div class="mc-label">Pokrivenost</div><div class="mc-val" style="color:{"var(--grn)" if total_days>90 else "var(--red)"}">{total_months_str} mes</div><div class="mc-sub">{total_days} dana</div></div>
          <div class="mag-card"><div class="mc-label">Artikala</div><div class="mc-val" style="color:var(--t2)">{len(df_mag)}</div><div class="mc-sub">u magacinu</div></div>
        </div>'''
        grupa_cards='<div class="mag-cards">'
        for g in sorted(mag_grupe_summary.keys()):
            gs=mag_grupe_summary[g]
            d=round(gs['kol']/(gs['avg']/30)) if gs['avg']>0 else 9999
            m_str=f"{d/30:.1f}" if d<9999 else "∞"
            clr='var(--grn)' if d>90 else ('var(--red)' if d<=30 else '#92400e')
            grupa_cards+=f'<div class="mag-card"><div class="mc-label">{g}</div><div class="mc-val" style="color:{clr}">{m_str} mes</div><div class="mc-sub">{fmtnum(gs["kol"])} kom · ø {fmtnum(gs["avg"])}/mes</div></div>'
        grupa_cards+='</div>'

        # Ucitaj JS i CSS iz originalnog koda
        JS_MESECNI = 'var PRODAJA_DATA=' + prodaja_json + ';\nvar PROFIT_DATA=' + profit_json + ';\nvar DRV_DATA=' + drv_json + ';\nvar NUM_MONTHS=' + str(num_s) + ';\nvar NUM_PROFIT_MONTHS=' + str(num_p) + ''';\nfunction showTab(n){document.querySelectorAll('.panel').forEach(function(p){p.classList.remove('active')});document.querySelectorAll('.tab').forEach(function(t){t.classList.remove('active')});document.getElementById('panel-'+n).classList.add('active');var ts=document.querySelectorAll('.tab');if(n==='prodaja')ts[0].classList.add('active');else if(n==='profit')ts[1].classList.add('active');else if(n==='drv')ts[2].classList.add('active');else ts[3].classList.add('active');document.getElementById('leg').style.display=(n==='prodaja')?'flex':'none';document.getElementById('filters-prodaja').style.display=(n==='prodaja')?'flex':'none';document.getElementById('filters-profit').style.display=(n==='profit')?'flex':'none';document.getElementById('filters-drv').style.display=(n==='drv')?'flex':'none';document.getElementById('filters-zalihe').style.display=(n==='zalihe')?'flex':'none'}
        function tog(id){var btn=document.getElementById('b-'+id);if(!btn)return;var isO=btn.textContent.trim()==='\u2212';var rows=document.querySelectorAll('tr[data-p="'+id+'"]');if(isO){rows.forEach(function(r){r.classList.add('hidden');var cb=r.querySelector('.beg');if(cb){cb.textContent='+';var cid=cb.id.replace('b-','');document.querySelectorAll('tr[data-p="'+cid+'"]').forEach(function(cr){cr.classList.add('hidden')})}});btn.textContent='+'}else{var fG=document.getElementById('f-grupa')?document.getElementById('f-grupa').value:'';rows.forEach(function(r){if(fG&&r.getAttribute('data-grupa')&&r.getAttribute('data-grupa')!==fG)return;r.classList.remove('hidden')});btn.textContent='\u2212'}}
        function toggleAll(o){var ap=document.querySelector('.panel.active');if(!ap)return;ap.querySelectorAll('.be').forEach(function(btn){var id=btn.id.replace('b-','');var rows=document.querySelectorAll('tr[data-p="'+id+'"]');rows.forEach(function(r){o?r.classList.remove('hidden'):r.classList.add('hidden')});btn.textContent=o?'\u2212':'+'})}
        function applyFilters(){var fS=document.getElementById('f-sistem').value;var fG=document.getElementById('f-grupa').value;var tbody=document.getElementById('tbody-prodaja');tbody.querySelectorAll('.be').forEach(function(btn){btn.textContent='+'});tbody.querySelectorAll('tr').forEach(function(r){var rS=r.getAttribute('data-sistem');var isSr=r.classList.contains('sr');var isGr=r.classList.contains('gr');var isAr=r.classList.contains('ar');var isSep=r.classList.contains('sep');var isTotal=r.classList.contains('totalrow');if(isTotal){r.classList.remove('hidden');return}if(fS&&rS&&rS!==fS){r.classList.add('hidden');return}if(isSr){if(fG){var hG=PRODAJA_DATA[rS]&&PRODAJA_DATA[rS][fG];if(!hG){r.classList.add('hidden');return}var cells=r.querySelectorAll('td');var gt=0;for(var i=0;i<NUM_MONTHS;i++){cells[i+1].textContent=hG[i].toLocaleString('sr-RS');gt+=hG[i]}cells[NUM_MONTHS+1].textContent=gt.toLocaleString('sr-RS')}else{var allG=PRODAJA_DATA[rS];if(allG){var cells=r.querySelectorAll('td');var sums=new Array(NUM_MONTHS).fill(0);for(var g in allG)for(var i=0;i<NUM_MONTHS;i++)sums[i]+=allG[g][i];var gt=0;for(var i=0;i<NUM_MONTHS;i++){cells[i+1].textContent=sums[i].toLocaleString('sr-RS');gt+=sums[i]}cells[NUM_MONTHS+1].textContent=gt.toLocaleString('sr-RS')}}r.classList.remove('hidden');return}if(isGr||isAr){r.classList.add('hidden');return}if(isSep){if(fS&&rS!==fS){r.classList.add('hidden');return}if(fG){var hSG=PRODAJA_DATA[rS]&&PRODAJA_DATA[rS][fG];if(!hSG){r.classList.add('hidden');return}}r.classList.remove('hidden');return}});recalcTotals(fS,fG)}
        function recalcTotals(fS,fG){var t=new Array(NUM_MONTHS).fill(0);for(var s in PRODAJA_DATA){if(fS&&s!==fS)continue;for(var g in PRODAJA_DATA[s]){if(fG&&g!==fG)continue;var v=PRODAJA_DATA[s][g];for(var i=0;i<NUM_MONTHS;i++)t[i]+=v[i]}}var gt=t.reduce(function(a,b){return a+b},0);var tr=document.getElementById('prodaja-total');if(!tr)return;var c=tr.querySelectorAll('td');for(var i=1;i<=NUM_MONTHS;i++)c[i].textContent=t[i-1].toLocaleString('sr-RS');c[NUM_MONTHS+1].textContent=gt.toLocaleString('sr-RS')}
        function resetFilters(){document.getElementById('f-sistem').value='';document.getElementById('f-grupa').value='';applyFilters()}
        function fmtN(v){return v.toLocaleString('sr-RS')}
        function applyProfitFilters(){var fS=document.getElementById('fp-sistem').value;var tbody=document.getElementById('tbody-profit');var checks=document.querySelectorAll('#filters-profit input[type=checkbox]');var ac={};checks.forEach(function(cb){ac[cb.value]=cb.checked});tbody.querySelectorAll('.be').forEach(function(btn){btn.textContent='+'});tbody.querySelectorAll('tr').forEach(function(r){var rS=r.getAttribute('data-sistem');var isSr=r.classList.contains('sr');var isSep=r.classList.contains('sep');var isTotal=r.classList.contains('totalrow');if(isTotal){r.classList.remove('hidden');return}if(isSr||isSep){if(fS&&rS&&rS!==fS)r.classList.add('hidden');else r.classList.remove('hidden');return}r.classList.add('hidden')});tbody.querySelectorAll('tr[data-trosak]').forEach(function(r){var tid=r.getAttribute('data-trosak');r.classList.toggle('excluded',!ac[tid])});var grandT=new Array(NUM_PROFIT_MONTHS).fill(0);tbody.querySelectorAll('.ukupni-row').forEach(function(ukRow){var sid=ukRow.getAttribute('data-sid');var sistem=ukRow.getAttribute('data-sistem');var pd=PROFIT_DATA[sistem];if(!pd)return;var uk=new Array(NUM_PROFIT_MONTHS).fill(0);if(ac['mkt'])for(var i=0;i<NUM_PROFIT_MONTHS;i++)uk[i]+=pd['mkt'][i];for(var tid in pd){if(tid==='profit'||tid==='mkt')continue;if(ac[tid])for(var i=0;i<NUM_PROFIT_MONTHS;i++)uk[i]+=pd[tid][i]}var ukT=uk.reduce(function(a,b){return a+b},0);var cells=ukRow.querySelectorAll('td');for(var i=3;i<3+NUM_PROFIT_MONTHS;i++)cells[i].textContent=fmtN(uk[i-3]);cells[3+NUM_PROFIT_MONTHS].textContent=fmtN(ukT);var nr=tbody.querySelector('.neto-row[data-sid="'+sid+'"]');if(!nr)return;var p=pd['profit'];var nc=nr.querySelectorAll('td');var nt=0;for(var i=0;i<NUM_PROFIT_MONTHS;i++){var nv=p[i]-uk[i];nt+=nv;nc[3+i].textContent=fmtN(nv);nc[3+i].className='n '+(nv>=0?'np':'nn')}nc[3+NUM_PROFIT_MONTHS].textContent=fmtN(nt);nc[3+NUM_PROFIT_MONTHS].className='n '+(nt>=0?'np':'nn');nc[3+NUM_PROFIT_MONTHS].style.fontSize='13px';var sr=tbody.querySelector('.profit-sr[data-sid="'+sid+'"]');if(sr){var sc=sr.querySelectorAll('td');for(var i=3;i<3+NUM_PROFIT_MONTHS;i++){var nv2=p[i-3]-uk[i-3];sc[i].textContent=fmtN(nv2);sc[i].className='nb '+(nv2>=0?'np':'nn')}var stot=p.reduce(function(a,b){return a+b},0)-ukT;sc[3+NUM_PROFIT_MONTHS].textContent=fmtN(stot);sc[3+NUM_PROFIT_MONTHS].className='nb '+(stot>=0?'np':'nn');sc[3+NUM_PROFIT_MONTHS].style.fontSize='13px'}if(!fS||sistem===fS){for(var i=0;i<NUM_PROFIT_MONTHS;i++)grandT[i]+=p[i]-uk[i]}});var tRow=document.getElementById('profit-total');if(tRow){var tc=tRow.querySelectorAll('td');var ggt=grandT.reduce(function(a,b){return a+b},0);for(var i=3;i<3+NUM_PROFIT_MONTHS;i++){tc[i].textContent=fmtN(grandT[i-3]);tc[i].className='nb total-cell '+(grandT[i-3]>=0?'np':'nn')}tc[3+NUM_PROFIT_MONTHS].textContent=fmtN(ggt);tc[3+NUM_PROFIT_MONTHS].className='nb total-cell '+(ggt>=0?'np':'nn');tc[3+NUM_PROFIT_MONTHS].style.fontSize='13px'}}
        function resetProfitFilters(){document.getElementById('fp-sistem').value='';document.querySelectorAll('#filters-profit input[type=checkbox]').forEach(function(cb){cb.checked=true});applyProfitFilters()}
        function applyDrvFilters(){var fS=document.getElementById('fd-sistem').value;var tbody=document.getElementById('tbody-drv');tbody.querySelectorAll('.be').forEach(function(btn){btn.textContent='+'});tbody.querySelectorAll('tr').forEach(function(r){var rS=r.getAttribute('data-sistem');var isSr=r.classList.contains('sr');var isSep=r.classList.contains('sep');var isTotal=r.classList.contains('totalrow');if(isTotal){r.classList.remove('hidden');return}if(isSr||isSep){if(fS&&rS&&rS!==fS)r.classList.add('hidden');else r.classList.remove('hidden');return}r.classList.add('hidden')});var grandT=new Array(NUM_PROFIT_MONTHS).fill(0);for(var s in DRV_DATA){if(fS&&s!==fS)continue;var d=DRV_DATA[s];for(var i=0;i<NUM_PROFIT_MONTHS;i++)grandT[i]+=d['profit'][i]-d['mkt'][i]}var tRow=document.getElementById('drv-total');if(tRow){var tc=tRow.querySelectorAll('td');var ggt=grandT.reduce(function(a,b){return a+b},0);for(var i=3;i<3+NUM_PROFIT_MONTHS;i++){tc[i].textContent=fmtN(grandT[i-3]);tc[i].className='nb total-cell '+(grandT[i-3]>=0?'np':'nn')}tc[3+NUM_PROFIT_MONTHS].textContent=fmtN(ggt);tc[3+NUM_PROFIT_MONTHS].className='nb total-cell '+(ggt>=0?'np':'nn');tc[3+NUM_PROFIT_MONTHS].style.fontSize='13px'}}
        function applyZaliheFilters(){var fS=document.getElementById('fz-sistem').value;var tbody=document.getElementById('tbody-zalihe');tbody.querySelectorAll('.be').forEach(function(btn){btn.textContent='+'});tbody.querySelectorAll('tr').forEach(function(r){var rS=r.getAttribute('data-sistem');var isSr=r.classList.contains('sr');var isSep=r.classList.contains('sep');var isTotal=r.classList.contains('totalrow');if(isTotal){r.classList.remove('hidden');return}if(isSr||isSep){if(fS&&rS&&rS!==fS)r.classList.add('hidden');else r.classList.remove('hidden');return}r.classList.add('hidden')})}'''

        CSS_MESECNI = '''*{margin:0;padding:0;box-sizing:border-box}
        :root{--bg:#f4f6fb;--bg2:#ffffff;--bd:#e2e6ef;--bd2:#d0d5e0;--t1:#1a1a2e;--t2:#5a5f7a;--t3:#8b90a5;--ac:#a855f7;--acd:rgba(168,85,247,0.08);--red:#ec4899;--redd:rgba(236,72,153,0.06);--grn:#7c3aed;--grnd:rgba(124,58,237,0.06);--shadow:0 1px 3px rgba(0,0,0,0.06)}
        body{font-family:'Outfit',sans-serif;background:var(--bg);color:var(--t1)}
        .hdr{padding:16px 24px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:12px;background:#12002a;border-bottom:3px solid;border-image:linear-gradient(90deg,#a855f7,#ec4899) 1}
        .hdr h1{font-size:18px;font-weight:700;color:white;letter-spacing:0.5px}
        .hdr .sub{font-size:11px;color:rgba(255,255,255,0.4);margin-top:2px}
        .badge{padding:4px 10px;border-radius:20px;font-size:10px;font-weight:600}
        .bg{background:rgba(124,58,237,0.15);color:#a855f7;border:1px solid rgba(168,85,247,0.3)}
        .br{background:rgba(236,72,153,0.12);color:#ec4899;border:1px solid rgba(236,72,153,0.3)}
        .toolbar{display:flex;gap:0;background:white;border-bottom:1px solid var(--bd)}
        .tabs{display:flex;gap:0;padding:0 16px;flex:1}
        .tab{padding:12px 20px;font-size:11px;font-weight:600;cursor:pointer;border-bottom:3px solid transparent;color:var(--t3);transition:all .2s;margin-bottom:-1px;user-select:none;font-family:monospace}
        .tab:hover{color:var(--t1)}.tab.active{color:#a855f7;border-bottom-color:#a855f7}
        .filters{display:flex;gap:12px;align-items:center;flex-wrap:wrap;padding:10px 20px;background:var(--bg);border-bottom:1px solid var(--bd)}
        .filters label{font-size:10px;color:var(--t3);font-weight:600;text-transform:uppercase;letter-spacing:.5px}
        .filters select{background:var(--bg2);color:var(--t1);border:1px solid var(--bd2);border-radius:8px;padding:5px 10px;font-size:12px;cursor:pointer;min-width:140px}
        .reset-btn{background:rgba(236,72,153,0.08);color:#ec4899;border:1px solid rgba(236,72,153,0.2);padding:5px 12px;border-radius:8px;cursor:pointer;font-size:11px;font-weight:600}
        .tcb{display:inline-flex;align-items:center;gap:4px;font-size:10px;color:var(--t2);cursor:pointer;padding:3px 8px;border-radius:6px;background:var(--bg2);border:1px solid var(--bd)}
        .tcb input{accent-color:#a855f7;cursor:pointer}
        .legend{display:flex;gap:8px;align-items:center;flex-wrap:wrap;padding:6px 20px;font-size:10px;color:var(--t3);background:var(--bg);border-bottom:1px solid var(--bd)}
        .legend b{font-weight:700}.lc{display:inline-block;padding:2px 7px;border-radius:4px;font-weight:700;font-size:9px;color:#1a1a2e}
        .panel{display:none}.panel.active{display:block}
        .tw{overflow-x:auto;padding:4px 12px 32px}
        table{border-collapse:collapse;width:100%;font-size:11px;margin-top:4px;background:var(--bg2);border-radius:10px;overflow:hidden;box-shadow:var(--shadow)}
        thead th{background:linear-gradient(180deg,#f8f9fd,#eef0f7);color:var(--t2);font-weight:700;font-size:10px;font-family:monospace;text-transform:uppercase;letter-spacing:.5px;padding:9px 7px;text-align:center;border-bottom:2px solid var(--bd);position:sticky;top:0;z-index:20;white-space:nowrap}
        td{padding:5px 7px;border-bottom:1px solid var(--bd);white-space:nowrap;font-family:monospace;font-size:11px}
        tr.sr{background:rgba(168,85,247,0.03)}tr.sr:hover{background:rgba(168,85,247,0.07)}tr.sr td{border-bottom:1px solid var(--bd2)}
        tr.gr{background:rgba(168,85,247,0.015)}tr.gr:hover{background:rgba(168,85,247,0.05)}
        tr.ar{background:var(--bg2)}tr.ar:hover{background:rgba(0,0,0,0.015)}
        tr.cr{background:rgba(236,72,153,0.02)}tr.ctr{background:rgba(236,72,153,0.04)}tr.nr{background:rgba(124,58,237,0.05)}
        tr.sep td{height:4px;background:var(--bg);border:none;padding:0}
        .be{display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;border-radius:6px;background:var(--acd);color:#a855f7;font-size:13px;font-weight:800;cursor:pointer;border:none;margin-right:8px;font-family:monospace;vertical-align:middle;transition:all .15s}
        .be:hover{background:#a855f7;color:#fff}.beg{background:rgba(90,95,122,0.06);color:var(--t2)}.beg:hover{background:var(--t2);color:#fff}
        .sn{font-weight:700;color:#a855f7;font-size:12px;cursor:pointer}.gn{color:var(--t2);font-weight:600;cursor:pointer}
        .an{padding-left:36px;color:var(--t3);font-size:10px}.an::before{content:'↳ ';color:var(--bd2)}
        .n{text-align:right}.nb{text-align:right;font-weight:600}.nt{text-align:right;font-weight:800;color:#a855f7}
        .cc{color:#ec4899;font-style:italic;opacity:.75}.cl{color:#ec4899;font-style:italic;font-size:10px;opacity:.75}
        .ctl{color:#be185d;font-weight:800;font-size:10px}.ctc{color:#be185d;font-weight:700}
        .nl{color:#7c3aed;font-weight:800;font-size:10px}.np{color:#7c3aed;font-weight:700}.nn{color:#ec4899;font-weight:700}
        .nf{color:#ec4899;font-weight:600;font-size:10px}.no{color:#a855f7;font-weight:600;font-size:10px}
        .hidden{display:none}
        tr.totalrow{background:linear-gradient(90deg,rgba(168,85,247,0.06),rgba(236,72,153,0.06))}
        tr.totalrow td{border-top:2px solid var(--bd2);padding:9px 7px}
        .total-label{font-weight:800;color:#a855f7;font-size:13px;font-family:monospace;letter-spacing:1px}
        .total-cell{font-size:12px}
        .mag-section{padding:16px;background:var(--bg2);border-radius:10px;margin:8px 12px;box-shadow:var(--shadow)}
        .mag-section h3{font-size:13px;font-weight:700;color:#a855f7;margin-bottom:4px}
        .mag-cards{display:flex;gap:12px;margin-bottom:16px;flex-wrap:wrap}
        .mag-card{flex:1;min-width:140px;padding:12px 16px;border-radius:10px;background:var(--bg);border:1px solid var(--bd)}
        .mag-card .mc-label{font-size:9px;text-transform:uppercase;letter-spacing:.5px;color:var(--t3);font-weight:600}
        .mag-card .mc-val{font-size:20px;font-weight:800;font-family:monospace;margin-top:4px}
        .mag-card .mc-sub{font-size:9px;color:var(--t3);margin-top:2px}
        .hb button{padding:5px 12px;border-radius:8px;cursor:pointer;font-size:10px;font-weight:600}'''

        html_out = f'''<!DOCTYPE html><html lang="sr"><head><meta charset="UTF-8">
        <link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600;700&family=Outfit:wght@300;400;600;700;800&display=swap" rel="stylesheet">
        <style>{CSS_MESECNI}</style></head><body>
        <div class="hdr"><div><h1>&#128202; Mesečni izveštaj prodaje</h1><div class="sub">{info}</div></div>
          <div class="hb" style="display:flex;gap:8px;align-items:center">{badge_html}
    <button onclick="toggleAll(true)" style="background:rgba(168,85,247,0.1);color:#a855f7;border:1px solid rgba(168,85,247,0.2)">Otvori sve</button>
    <button onclick="toggleAll(false)" style="background:rgba(90,95,122,0.06);color:var(--t2);border:1px solid var(--bd2)">Zatvori sve</button></div></div>
        <div class="toolbar"><div class="tabs">
          <div class="tab active" onclick="showTab('prodaja')">📊 PRODAJA</div>
          <div class="tab" onclick="showTab('profit')">💰 PROFITABILNOST</div>
          <div class="tab" onclick="showTab('drv')">DR VUKAŠIN</div>
          <div class="tab" onclick="showTab('zalihe')">📦 ZALIHE</div>
        </div></div>
        <div class="filters" id="filters-prodaja"><label>Sistem:</label><select id="f-sistem" onchange="applyFilters()"><option value="">Svi</option>{sistem_options}</select><label>Grupa:</label><select id="f-grupa" onchange="applyFilters()"><option value="">Sve</option>{grupa_options}</select><button class="reset-btn" onclick="resetFilters()">✕ Reset</button></div>
        <div class="filters hidden" id="filters-profit"><label>Sistem:</label><select id="fp-sistem" onchange="applyProfitFilters()"><option value="">Svi</option>{sistem_options}</select>&nbsp;<label>Troškovi:</label>{trosak_checks}<button class="reset-btn" onclick="resetProfitFilters()">✕ Reset</button></div>
        <div class="filters hidden" id="filters-drv"><label>Sistem:</label><select id="fd-sistem" onchange="applyDrvFilters()"><option value="">Svi</option>{sistem_options}</select></div>
        <div class="filters hidden" id="filters-zalihe"><label>Sistem:</label><select id="fz-sistem" onchange="applyZaliheFilters()"><option value="">Svi</option>{sistem_options}</select></div>
        <div class="legend" id="leg"><b>NERD:</b><span class="lc" style="background:#90EE90">1390</span><span class="lc" style="background:#FFD1DC">1300</span><span class="lc" style="background:#FFB6C1">1290</span><span class="lc" style="background:#FF69B4;color:#fff">1190</span><span class="lc" style="background:#C71585;color:#fff">990</span>&nbsp;<b>HQD:</b><span class="lc" style="background:#90EE90">890</span><span class="lc" style="background:#FFD1DC">800</span><span class="lc" style="background:#FFB6C1">790</span><span class="lc" style="background:#FF69B4;color:#fff">730</span><span class="lc" style="background:#C71585;color:#fff">690</span></div>
        <div class="panel active" id="panel-prodaja"><div class="tw"><table><thead>{ph}</thead><tbody id="tbody-prodaja">{"".join(prodaja_rows)}</tbody></table></div></div>
        <div class="panel" id="panel-profit"><div class="tw"><table><thead>{prh}</thead><tbody id="tbody-profit">{"".join(profit_rows)}</tbody></table></div></div>
        <div class="panel" id="panel-drv"><div class="tw"><table><thead>{prh}</thead><tbody id="tbody-drv">{"".join(drv_rows)}</tbody></table></div></div>
        <div class="panel" id="panel-zalihe">
        <div class="mag-section"><h3>STANJE MAGACINA</h3><div style="font-size:10px;color:var(--t3);margin-bottom:12px">{mag_info}</div>{cards_html}{grupa_cards}<div class="tw" style="padding:0"><table><thead>{mh}</thead><tbody>{"".join(mag_rows)}</tbody></table></div></div>
        <div style="padding:16px 20px 8px;font-size:12px;font-weight:700;color:var(--t2)">ZALIHE PO SISTEMIMA</div>
        <div class="tw"><table><thead>{zsh}</thead><tbody id="tbody-zalihe">{"".join(zal_rows)}</tbody></table></div></div>
        <script>{JS_MESECNI}</script></body></html>'''

        return html_out


    with st.spinner("⏳ Učitavam podatke..."):
        html_content = build_mesecni_html()

    if html_content is None:
        st.error("❌ Podaci nisu dostupni. Proveri da li su fajlovi postavljeni na GitHub.")
    else:
        components.html(html_content, height=900, scrolling=True)

elif page == 'finansijski':
    render_header("Finansijski izveštaj · Dugovanja · Lager")

    @st.cache_data(ttl=300)
    def build_finansijski_html():
        import json as json_mod, numpy as np_f

        buf_s = load_github_excel("tabela sistemi3.xlsx")
        if buf_s is None:
            return None
        cfg = load_github_config()
        ukljuci_poslednji = cfg.get("ukljuci_poslednji_mesec", False)

        df = pd.read_excel(buf_s, sheet_name='tabela')
        df.columns = df.columns.astype(str).str.strip()
        df = df[df['SISTEM'] != 'MOBILLAND']

        num_cols = ['Godina','Mesec','VREDNOST PROMETA','VREDNOST PRODAJA KA KRAJNJEM KUPCU',
                    'MESECNI TROSAK1','KNJIZNO TOTAL','DODATNI MESECNI TROSAK','PROFIT2','PROFIT3',
                    'VREDNOST LAGERA','VREDNOST LAGERA NC','Promet','Prodata kolicina ka krajnjem kupcu',
                    'Zalihe','POVRAT STARIH CIGARETA sa pdv','POVRAT NOVIH CIGARETA sa PDV',
                    'POCETNO STANJE','POVRAT TOTAL sa PDV','UPLATA','UPLATA KA SISTEMIMA','UPLATA KA SISTEMU 1',
                    'FINALNA MP','NJIHOVA ZARADA 1','Nacin placanja','VP CENA SA PDVOM','KNJIZNO PROVERA',
                    'Fakturisano sa njihove strane']
        for c in num_cols:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)

        nacin_map = {}
        dfz = df[df['Mesec'] == 0]
        for s in dfz['SISTEM'].unique():
            if pd.notna(s):
                nacin_map[str(s)] = int(dfz[dfz['SISTEM']==s]['Nacin placanja'].iloc[0])

        brendovi = sorted([str(g) for g in df['Grupa artikla'].unique() if pd.notna(g) and str(g).strip() != ''])
        for b in brendovi:
            df[f'Q_{b}'] = df.apply(lambda x: x['Prodata kolicina ka krajnjem kupcu'] if str(x['Grupa artikla']) == b else 0, axis=1)
        df['UPLATA_KA_NJIMA_TOTAL'] = df['UPLATA KA SISTEMIMA'].fillna(0) + df['UPLATA KA SISTEMU 1'].fillna(0)
        df['TOTAL_TROSAK'] = df['MESECNI TROSAK1'].fillna(0) + df['KNJIZNO TOTAL'].fillna(0) + df['DODATNI MESECNI TROSAK'].fillna(0)

        agg_dict = {
            'VREDNOST PROMETA':'sum','VREDNOST PRODAJA KA KRAJNJEM KUPCU':'sum',
            'MESECNI TROSAK1':'sum','KNJIZNO TOTAL':'sum','DODATNI MESECNI TROSAK':'sum','TOTAL_TROSAK':'sum',
            'PROFIT2':'sum','PROFIT3':'sum','VREDNOST LAGERA':'sum','VREDNOST LAGERA NC':'sum',
            'Promet':'sum','Prodata kolicina ka krajnjem kupcu':'sum','Zalihe':'sum',
            'POVRAT STARIH CIGARETA sa pdv':'sum','POVRAT NOVIH CIGARETA sa PDV':'sum',
            'POCETNO STANJE':'sum','POVRAT TOTAL sa PDV':'sum','UPLATA':'sum','UPLATA_KA_NJIMA_TOTAL':'sum',
            'NJIHOVA ZARADA 1':'sum','KNJIZNO PROVERA':'sum','Fakturisano sa njihove strane':'sum',
        }
        for b in brendovi: agg_dict[f'Q_{b}'] = 'sum'

        grouped = df.groupby(['SISTEM','Godina','Mesec']).agg(agg_dict).reset_index()
        grouped = grouped.replace([np_f.inf, -np_f.inf], 0).fillna(0)

        all_ym = grouped[grouped['Mesec'] > 0][['Godina','Mesec']].drop_duplicates()
        all_ym['ym_key'] = all_ym['Godina'] * 100 + all_ym['Mesec']
        max_ym = all_ym['ym_key'].max()
        max_godina = int(max_ym // 100)
        max_mesec = int(max_ym % 100)

        if not ukljuci_poslednji:
            ym_sorted = sorted(all_ym['ym_key'].unique())
            target_ym = ym_sorted[-2] if len(ym_sorted) >= 2 else ym_sorted[-1]
        else:
            target_ym = max_ym

        lager_mesec = int(target_ym % 100); lager_godina = int(target_ym // 100)

        lager_grupe = {}
        for sistem in df['SISTEM'].unique():
            if pd.isna(sistem): continue
            s = str(sistem)
            ds = df[(df['SISTEM']==sistem) & (df['Mesec']>0)].copy()
            if ds.empty: continue
            ds['ym'] = ds['Godina']*100 + ds['Mesec']
            dl = ds[(ds['Godina']==lager_godina) & (ds['Mesec']==lager_mesec)]
            if dl.empty:
                sys_last = ds['ym'].max()
                dl = ds[ds['ym']==sys_last]
                lm = int(sys_last % 100); lg = int(sys_last // 100)
            else:
                lm = lager_mesec; lg = lager_godina
            grupe = {}
            for _, row in dl.iterrows():
                g = str(row['Grupa artikla']) if pd.notna(row['Grupa artikla']) else ''
                if g.strip() == '': continue
                if g not in grupe: grupe[g] = {'lager': 0, 'vp_cena_pdv': 0}
                grupe[g]['lager'] += int(row['Zalihe'])
                grupe[g]['vp_cena_pdv'] = round(float(row['VP CENA SA PDVOM']), 2)
            if grupe:
                grupe_list = []
                for k, v in sorted(grupe.items()):
                    vrednost = round(v['lager'] * v['vp_cena_pdv'], 2)
                    grupe_list.append({'grupa': k, 'lager': v['lager'], 'vp_cena_pdv': v['vp_cena_pdv'], 'vrednost': vrednost})
                lager_grupe[s] = {'mesec': lm, 'godina': lg, 'grupe': grupe_list}

        if not ukljuci_poslednji:
            grouped = grouped[~((grouped['Godina'] == max_godina) & (grouped['Mesec'] == max_mesec))]

        records = []
        for _, row in grouped.iterrows():
            r = {
                'sistem':str(row['SISTEM']),'godina':int(row['Godina']),'mesec':int(row['Mesec']),
                'v_promet':round(float(row['VREDNOST PROMETA']),2),
                'v_kupac':round(float(row['VREDNOST PRODAJA KA KRAJNJEM KUPCU']),2),
                'marketing':round(float(row['MESECNI TROSAK1']),2),
                'knjizno':round(float(row['KNJIZNO TOTAL']),2),
                'dodatni_trosak':round(float(row['DODATNI MESECNI TROSAK']),2),
                'total_trosak':round(float(row['TOTAL_TROSAK']),2),
                'profit_prodaja':round(float(row['PROFIT2']),2),
                'profit_promet':round(float(row['PROFIT3']),2),
                'v_lager_vp':round(float(row['VREDNOST LAGERA']),2),
                'v_lager_nc':round(float(row['VREDNOST LAGERA NC']),2),
                'q_promet':round(float(row['Promet']),2),
                'q_kupac':round(float(row['Prodata kolicina ka krajnjem kupcu']),2),
                'q_lager':round(float(row['Zalihe']),2),
                'pov_stari':round(float(row['POVRAT STARIH CIGARETA sa pdv']),2),
                'pov_novi':round(float(row['POVRAT NOVIH CIGARETA sa PDV']),2),
                'poc_stanje':round(float(row['POCETNO STANJE']),2),
                'pov_total':round(float(row['POVRAT TOTAL sa PDV']),2),
                'uplata':round(float(row['UPLATA']),2),
                'uplata_ka_njima':round(float(row['UPLATA_KA_NJIMA_TOTAL']),2),
                'njihova_zarada_1':round(float(row['NJIHOVA ZARADA 1']),2),
                'knjizno_provera':round(float(row['KNJIZNO PROVERA']),2),
                'fakturisano':round(float(row['Fakturisano sa njihove strane']),2),
            }
            for b in brendovi: r[f'q_{b}'] = round(float(row[f'Q_{b}']),2)
            records.append(r)

        prodaja_grupe = {}
        for sistem in df['SISTEM'].unique():
            if pd.isna(sistem): continue
            s = str(sistem)
            ds = df[(df['SISTEM']==sistem) & (df['Mesec']>0)].copy()
            if ds.empty: continue
            ds['ym'] = (ds['Godina']*100 + ds['Mesec']).astype(int)
            periods = sorted(ds['ym'].unique())
            grupe_pg = sorted([str(g) for g in ds['Grupa artikla'].unique() if pd.notna(g) and str(g).strip() != ''])
            if not grupe_pg: continue
            data_pg = {}
            for g in grupe_pg:
                data_pg[g] = {}
                for ym in periods:
                    dg = ds[(ds['Grupa artikla'].astype(str)==g) & (ds['ym']==ym)]
                    qty = int(dg['Prodata kolicina ka krajnjem kupcu'].sum())
                    mp = float(dg['FINALNA MP'].mean()) if not dg.empty and dg['FINALNA MP'].sum()!=0 else 0
                    data_pg[g][str(int(ym))] = {'qty': qty, 'mp': round(mp)}
            prodaja_grupe[s] = {'periods': [int(p) for p in periods], 'grupe': grupe_pg, 'data': data_pg}

        meta = {
            'sistemi': sorted(grouped['SISTEM'].unique().tolist()),
            'godine': sorted([int(x) for x in grouped['Godina'].unique()]),
            'brendovi': brendovi,
            'ukljucen_poslednji': ukljuci_poslednji,
            'nacin_placanja': nacin_map,
            'lager_grupe': lager_grupe,
            'prodaja_grupe': prodaja_grupe,
        }
        data_json = json_mod.dumps({'meta': meta, 'data': records}, ensure_ascii=False)

        # Ucitaj originalni React HTML ali zameni boje u ljubicastu temu
        html_out = '''<!DOCTYPE html>
<html lang="sr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Finansijski Izvestaj VAPE</title>
<script src="https://unpkg.com/react@18/umd/react.production.min.js"></''' + '''script>
<script src="https://unpkg.com/react-dom@18/umd/react-dom.production.min.js"></''' + '''script>
<script src="https://unpkg.com/@babel/standalone/babel.min.js"></''' + '''script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></''' + '''script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.2/jspdf.plugin.autotable.min.js"></''' + '''script>
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600;700&display=swap" rel="stylesheet"/>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{background:#12002a;font-family:'DM Sans',-apple-system,sans-serif}
select{padding:7px 10px;border:1px solid rgba(168,85,247,0.3);border-radius:6px;background:rgba(255,255,255,0.06);color:#e2e6ed;font-size:13px;font-family:inherit;outline:none;cursor:pointer}
select:hover{border-color:#a855f7}
table{width:100%;border-collapse:collapse}
::-webkit-scrollbar{height:6px;width:6px}
::-webkit-scrollbar-track{background:#12002a}
::-webkit-scrollbar-thumb{background:rgba(168,85,247,0.3);border-radius:3px}
.ychk{display:flex;gap:12px;align-items:center}
.ychk label{display:flex;align-items:center;gap:5px;color:#c9d1d9;font-size:13px;cursor:pointer}
.ychk input[type=checkbox]{accent-color:#a855f7;width:16px;height:16px;cursor:pointer}
</style>
</head>
<body>
<div id="root"></div>
<script type="text/babel">
const {useState,useMemo}=React;
const RAW=''' + data_json + ''';
const META=RAW.meta,DATA=RAW.data.filter(r=>r.mesec>0);
const ML={1:'Jan',2:'Feb',3:'Mar',4:'Apr',5:'Maj',6:'Jun',7:'Jul',8:'Avg',9:'Sep',10:'Okt',11:'Nov',12:'Dec'};
const HAS_DOD=DATA.some(r=>(r.dodatni_trosak||0)!==0);
const fmt=v=>{if(v===null||v===undefined||v==='')return'-';const n=Number(v);if(isNaN(n))return'-';if(Math.abs(n)<0.5)return'-';const s=Math.round(Math.abs(n)).toLocaleString('sr-Latn-RS');return n<0?'('+s+')':s};
const fmtD=(v,d=1)=>{if(v==null)return'-';const n=Number(v);if(isNaN(n)||!isFinite(n))return'-';return n.toFixed(d)};
const fmtPdf=(v)=>{if(v===null||v===undefined)return'-';const n=Number(v);if(isNaN(n))return'-';if(Math.abs(n)<0.5)return'-';return Math.round(Math.abs(n)).toLocaleString('sr-Latn-RS')};
const fmtPdf2=(v)=>{if(v===null||v===undefined)return'-';const n=Number(v);if(isNaN(n))return'-';return n.toLocaleString('sr-Latn-RS',{minimumFractionDigits:2,maximumFractionDigits:2})};
const calcNjZarada=(rows,idx)=>rows[idx].njihova_zarada_1||0;
const calcNjZaradaTotal=(rows)=>{let t=0;for(let i=0;i<rows.length;i++)t+=calcNjZarada(rows,i);return t};
const TabBtn=({active,onClick,children})=><button onClick={onClick} style={{padding:'10px 28px',border:'none',borderBottom:active?'2px solid #a855f7':'2px solid transparent',background:'transparent',color:active?'#e2e6ed':'#6b7280',fontWeight:active?700:500,fontSize:14,cursor:'pointer',fontFamily:'inherit',transition:'all 0.2s'}}>{children}</button>;
const TH=({children,colSpan,style})=><th colSpan={colSpan} style={{padding:'8px 12px',background:'rgba(168,85,247,0.08)',color:'#94a3b8',fontSize:11,fontWeight:600,textTransform:'uppercase',letterSpacing:'0.05em',borderBottom:'1px solid rgba(168,85,247,0.2)',position:'sticky',top:0,zIndex:2,whiteSpace:'nowrap',...style}}>{children}</th>;
const TG=({children,colSpan})=><th colSpan={colSpan} style={{padding:'6px 12px',background:'rgba(168,85,247,0.05)',color:'#a855f7',fontSize:10,fontWeight:700,textTransform:'uppercase',letterSpacing:'0.08em',borderBottom:'2px solid rgba(168,85,247,0.3)',textAlign:'center'}}>{children}</th>;
const TD=({children,style,neg})=>{const isN=typeof children==='string'&&children.startsWith('(');return<td style={{padding:'7px 12px',borderBottom:'1px solid rgba(168,85,247,0.08)',fontSize:13,textAlign:'right',fontVariantNumeric:'tabular-nums',color:(neg&&isN)?'#ec4899':(neg&&children!=='-'&&!isN)?'#a855f7':'#c9d1d9',fontFamily:"'JetBrains Mono',monospace",...style}}>{children}</td>};
const TDL=({children})=><td style={{padding:'7px 12px',borderBottom:'1px solid rgba(168,85,247,0.08)',fontSize:13,fontWeight:600,color:'#e2e6ed',position:'sticky',left:0,background:'#1a0533',zIndex:1,whiteSpace:'nowrap'}}>{children}</td>;
const TC=({children,color})=><td style={{padding:'7px 12px',textAlign:'right',fontWeight:700,color:color||'#a855f7',fontSize:13,fontFamily:"'JetBrains Mono',monospace"}}>{children}</td>;
const Lbl=({children})=><span style={{fontSize:11,fontWeight:600,color:'#8a919e',textTransform:'uppercase',letterSpacing:'0.06em'}}>{children}</span>;
const SecH=({children,extra})=><div style={{padding:'10px 16px',background:'rgba(168,85,247,0.08)',borderRadius:'8px 8px 0 0',border:'1px solid rgba(168,85,247,0.2)',borderBottom:'2px solid rgba(168,85,247,0.4)',display:'flex',justifyContent:'space-between',alignItems:'center'}}><span style={{fontSize:11,fontWeight:700,color:'#a855f7',textTransform:'uppercase',letterSpacing:'0.08em'}}>{children}</span>{extra}</div>;
const DugCell=({v})=><td style={{padding:'7px 12px',textAlign:'right',fontWeight:700,fontSize:14,fontFamily:"'JetBrains Mono',monospace",color:v>0?'#ec4899':'#a855f7',background:v>0?'rgba(236,72,153,0.08)':'rgba(168,85,247,0.08)'}}>{fmt(v)}</td>;
''' + '''
function generateDebtPDF(sistem, dugF) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF('p', 'mm', 'a4');
  const pw = 210, ph = 297, ml = 22, mr = 22, cw = pw - ml - mr;
  const today = new Date();
  const mesNaz = ['januar','februar','mart','april','maj','jun','jul','avgust','septembar','oktobar','novembar','decembar'];
  const dateStr = today.getDate() + '. ' + mesNaz[today.getMonth()] + ' ' + today.getFullYear() + '.';
  const navy = [75, 0, 130]; const teal = [168, 85, 247]; const dark = [33, 37, 41];
  const mid = [108, 117, 125]; const light = [245, 240, 255]; const brd = [222, 210, 240];
  doc.setFillColor(...teal); doc.rect(0, 0, 5, ph, 'F');
  let y = 26;
  doc.setFontSize(26);doc.setFont('helvetica','bold');doc.setTextColor(...navy);
  doc.text('Finansijski pregled', ml, y); y += 10;
  doc.setFontSize(14);doc.setFont('helvetica','normal');doc.setTextColor(...teal);
  doc.text(sistem, ml, y);
  doc.setFontSize(10);doc.setTextColor(...mid); doc.text(dateStr, pw - mr, 30, {align:'right'});
  y += 8; doc.setDrawColor(...teal);doc.setLineWidth(0.6);doc.line(ml, y, ml+45, y); y += 16;
  const lg = META.lager_grupe[sistem];
  if (lg && lg.grupe && lg.grupe.length > 0) {
    doc.setFontSize(12);doc.setFont('helvetica','bold');doc.setTextColor(...navy);
    doc.text('PREGLED LAGERA', ml, y);
    doc.setFontSize(9);doc.setFont('helvetica','normal');doc.setTextColor(...mid);
    doc.text('Stanje: '+ML[lg.mesec]+' '+lg.godina+'.', ml+56, y); y += 8;
    const tblData = lg.grupe.map(g=>[g.grupa,g.lager.toLocaleString('sr-Latn-RS'),fmtPdf2(g.vp_cena_pdv),fmtPdf(g.vrednost)+' RSD']);
    const totL = lg.grupe.reduce((s,g)=>s+g.lager,0);
    const totV = lg.grupe.reduce((s,g)=>s+g.vrednost,0);
    doc.autoTable({startY:y,head:[['Grupa','Kolicina','VP cena (PDV)','Vrednost (PDV)']],body:tblData,foot:[['UKUPNO',totL.toLocaleString('sr-Latn-RS'),'',fmtPdf(totV)+' RSD']],theme:'plain',styles:{fontSize:10,cellPadding:{top:5,bottom:5,left:8,right:8},font:'helvetica',textColor:dark,lineWidth:0},headStyles:{fillColor:navy,textColor:[255,255,255],fontStyle:'bold',fontSize:9},footStyles:{fillColor:light,textColor:navy,fontStyle:'bold',fontSize:11},columnStyles:{0:{halign:'left',cellWidth:42},1:{halign:'right',cellWidth:28},2:{halign:'right',cellWidth:42},3:{halign:'right',cellWidth:52,fontStyle:'bold'}},margin:{left:ml,right:mr}});
    y = doc.lastAutoTable.finalY + 22;
    const vrLag = totV, zaUpl = dugF - vrLag;
    doc.setFontSize(12);doc.setFont('helvetica','bold');doc.setTextColor(...navy);
    doc.text('OBRACUN DUGA', ml, y); y += 10;
    const rH=18,lx=ml+6,vx=ml+cw-6;
    doc.setFillColor(...light);doc.rect(ml,y,cw,rH,'F');
    doc.setFontSize(10);doc.setFont('helvetica','normal');doc.setTextColor(...mid);
    doc.text('Dug po fakturi',lx,y+12);
    doc.setFontSize(14);doc.setFont('helvetica','bold');doc.setTextColor(...dark);
    doc.text(fmtPdf(dugF)+' RSD',vx,y+12,{align:'right'}); y+=rH+1;
    doc.setFontSize(10);doc.setFont('helvetica','normal');doc.setTextColor(...mid);
    doc.text('Vrednost lagera (sa PDV)',lx,y+12);
    doc.setFontSize(14);doc.setFont('helvetica','bold');doc.setTextColor(...dark);
    doc.text('- '+fmtPdf(vrLag)+' RSD',vx,y+12,{align:'right'}); y+=rH+3;
    doc.setFillColor(...navy);doc.roundedRect(ml,y,cw,26,2,2,'F');
    doc.setFontSize(11);doc.setFont('helvetica','bold');doc.setTextColor(200,180,240);
    doc.text('ZA UPLATU',lx,y+16);
    doc.setFontSize(18);doc.setFont('helvetica','bold');doc.setTextColor(255,255,255);
    doc.text((zaUpl>0?fmtPdf(zaUpl):'0')+' RSD',vx,y+17,{align:'right'});
  }
  doc.save('VREDNOST_DUGA_'+sistem.replace(/\\s+/g,'_')+'.pdf');
}
function calcDug(sistem){
  const allRows=DATA.filter(r=>r.sistem===sistem);
  const pocSt=RAW.data.filter(r=>r.sistem===sistem&&r.mesec===0).reduce((a,r)=>a+r.poc_stanje,0);
  const vProm=allRows.reduce((a,r)=>a+r.v_promet,0)*1.2;
  const pov=allRows.reduce((a,r)=>a+r.pov_total,0);
  const tros=allRows.reduce((a,r)=>a+r.knjizno,0)*1.2+allRows.reduce((a,r)=>a+(r.fakturisano||0),0);
  const upl=allRows.reduce((a,r)=>a+r.uplata,0);
  const uplN=allRows.reduce((a,r)=>a+r.uplata_ka_njima,0);
  const dugF=(pocSt+vProm)-(pov+tros+upl+uplN);
  const np=META.nacin_placanja[sistem]||0;
  let dugO=null;
  if(np===0){let lVP=0;for(let i=allRows.length-1;i>=0;i--)if(allRows[i].v_lager_vp!==0){lVP=allRows[i].v_lager_vp;break};dugO=dugF-lVP*1.2}
  return {dugF,dugO,np};
}
function aggBySistemMulti(rows,lastRows){const m={};const mSets={};rows.forEach(r=>{if(!m[r.sistem]){m[r.sistem]={sistem:r.sistem,v_promet:0,v_kupac:0,marketing:0,knjizno:0,dodatni_trosak:0,total_trosak:0,profit_prodaja:0,profit_promet:0,v_lager_vp:0,v_lager_nc:0,q_promet:0,q_kupac:0,q_lager:0,pov_stari:0,pov_novi:0,njihova_zarada_1:0,nj_zarada_calc:0,nM:0};mSets[r.sistem]=new Set()}const s=m[r.sistem];s.v_promet+=r.v_promet;s.v_kupac+=r.v_kupac;s.marketing+=r.marketing;s.knjizno+=r.knjizno;s.dodatni_trosak+=(r.dodatni_trosak||0);s.total_trosak+=r.total_trosak;s.profit_prodaja+=r.profit_prodaja;s.profit_promet+=r.profit_promet;s.q_promet+=r.q_promet;s.q_kupac+=r.q_kupac;s.pov_stari+=r.pov_stari;s.pov_novi+=r.pov_novi;s.njihova_zarada_1+=(r.njihova_zarada_1||0);mSets[r.sistem].add(r.godina*100+r.mesec)});Object.keys(m).forEach(s=>{m[s].nM=mSets[s].size});Object.keys(m).forEach(s=>{const sr=rows.filter(r=>r.sistem===s).sort((a,b)=>(a.godina*100+a.mesec)-(b.godina*100+b.mesec));m[s].nj_zarada_calc=calcNjZaradaTotal(sr)});lastRows.forEach(r=>{if(m[r.sistem]){m[r.sistem].v_lager_vp=r.v_lager_vp;m[r.sistem].v_lager_nc=r.v_lager_nc;m[r.sistem].q_lager=r.q_lager}});return Object.values(m).sort((a,b)=>a.sistem.localeCompare(b.sistem))}
function TotalPregled(){
  const allYears=META.godine.filter(y=>y>0);
  const [selYears,setSelYears]=useState(allYears.map(String));
  const [mOd,setMOd]=useState('1');const [mDo,setMDo]=useState('12');
  const mo=Number(mOd),md=Number(mDo);
  const toggleYear=y=>{const ys=String(y);setSelYears(prev=>prev.includes(ys)?prev.filter(x=>x!==ys):[...prev,ys])};
  const filt=useMemo(()=>{const ySet=new Set(selYears.map(Number));return DATA.filter(r=>ySet.has(r.godina)&&r.mesec>=mo&&r.mesec<=md)},[selYears,mo,md]);
  const lastRows=useMemo(()=>{if(!filt.length)return[];let mx=0;filt.forEach(r=>{const ym=r.godina*100+r.mesec;if(ym>mx)mx=ym});return DATA.filter(r=>r.godina===Math.floor(mx/100)&&r.mesec===mx%100)},[filt]);
  const agg=useMemo(()=>aggBySistemMulti(filt,lastRows),[filt,lastRows]);
  const tot=useMemo(()=>{const t={v_promet:0,v_kupac:0,marketing:0,knjizno:0,dodatni_trosak:0,total_trosak:0,profit_prodaja:0,profit_promet:0,v_lager_vp:0,v_lager_nc:0,pov_stari:0,pov_novi:0,njihova_zarada_1:0};agg.forEach(r=>Object.keys(t).forEach(k=>t[k]+=r[k]));return t},[agg]);
  const mOpts=Array.from({length:12},(_,i)=>({v:String(i+1),l:ML[i+1]}));
  return(<div style={{padding:'20px'}}>
    <div style={{display:'flex',gap:16,alignItems:'flex-end',marginBottom:20,flexWrap:'wrap'}}>
      <div style={{display:'flex',flexDirection:'column',gap:4}}><Lbl>Godine</Lbl><div className="ychk">{allYears.map(y=><label key={y}><input type="checkbox" checked={selYears.includes(String(y))} onChange={()=>toggleYear(y)}/>{y}</label>)}</div></div>
      <div style={{display:'flex',flexDirection:'column',gap:4}}><Lbl>Od meseca</Lbl><select value={mOd} onChange={e=>setMOd(e.target.value)} style={{width:130}}>{mOpts.map(o=><option key={o.v} value={o.v}>{o.l}</option>)}</select></div>
      <div style={{display:'flex',flexDirection:'column',gap:4}}><Lbl>Do meseca</Lbl><select value={mDo} onChange={e=>setMDo(e.target.value)} style={{width:130}}>{mOpts.map(o=><option key={o.v} value={o.v}>{o.l}</option>)}</select></div>
    </div>
    <div style={{overflowX:'auto',borderRadius:8,border:'1px solid rgba(168,85,247,0.2)'}}>
      <table style={{minWidth:1400}}>
        <thead>
          <tr><TG colSpan={2}>&nbsp;</TG><TG colSpan={2}>Promet (bez PDV)</TG><TG colSpan={HAS_DOD?4:3}>Trosak (bez PDV)</TG><TG colSpan={2}>Profit</TG><TG colSpan={2}>Vrednost Lagera</TG><TG colSpan={1}>Obrt</TG><TG colSpan={2}>Povrat</TG><TG colSpan={1}>Zarada</TG></tr>
          <tr><TH>Sistem</TH><TH>Placanje</TH><TH>Ka sistemu</TH><TH>Ka kupcu</TH><TH>Marketing</TH><TH>Knjizno</TH>{HAS_DOD&&<TH>Dodatni</TH>}<TH>Total</TH><TH>Po prodaji</TH><TH>Po prometu</TH><TH>VP cena</TH><TH>NC cena</TH><TH>Obrt lagera</TH><TH>Pov. starih</TH><TH>Pov. novih</TH><TH>Njihova</TH></tr>
        </thead>
        <tbody>
          {agg.map(r=>{const avgV=r.nM>0?r.v_kupac/r.nM:0;const obrt=avgV>0?r.v_lager_vp/avgV:0;const np=META.nacin_placanja[r.sistem];return(
            <tr key={r.sistem} onMouseEnter={e=>e.currentTarget.style.background='rgba(168,85,247,0.05)'} onMouseLeave={e=>e.currentTarget.style.background='transparent'}>
              <TDL>{r.sistem}</TDL>
              <TD style={{textAlign:'center',fontSize:11,color:np===1?'#a855f7':'#ec4899'}}>{np===1?'Faktura':'Odjava'}</TD>
              <TD>{fmt(r.v_promet)}</TD><TD>{fmt(r.v_kupac)}</TD><TD>{fmt(r.marketing)}</TD><TD>{fmt(r.knjizno)}</TD>{HAS_DOD&&<TD>{fmt(r.dodatni_trosak)}</TD>}<TD>{fmt(r.total_trosak)}</TD><TD neg>{fmt(r.profit_prodaja)}</TD><TD neg>{fmt(r.profit_promet)}</TD><TD>{fmt(r.v_lager_vp)}</TD><TD>{fmt(r.v_lager_nc)}</TD><TD>{fmtD(obrt)}</TD><TD>{fmt(r.pov_stari/1.2)}</TD><TD>{fmt(r.pov_novi/1.2)}</TD><TD>{fmt(r.nj_zarada_calc)}</TD>
            </tr>)})}
          <tr style={{background:'rgba(168,85,247,0.08)'}}>
            <td style={{padding:'9px 12px',fontWeight:700,color:'#a855f7',fontSize:13,position:'sticky',left:0,background:'rgba(168,85,247,0.08)',zIndex:1}}>GRAND TOTAL</td><td></td>
            {[tot.v_promet,tot.v_kupac,tot.marketing,tot.knjizno].map((v,i)=><TC key={i}>{fmt(v)}</TC>)}
            {HAS_DOD&&<TC>{fmt(tot.dodatni_trosak)}</TC>}
            <TC>{fmt(tot.total_trosak)}</TC>
            <TC color={tot.profit_prodaja>=0?'#a855f7':'#ec4899'}>{fmt(tot.profit_prodaja)}</TC>
            <TC color={tot.profit_promet>=0?'#a855f7':'#ec4899'}>{fmt(tot.profit_promet)}</TC>
            {[tot.v_lager_vp,tot.v_lager_nc].map((v,i)=><TC key={i}>{fmt(v)}</TC>)}
            <td style={{textAlign:'center',color:'#475569',fontSize:11}}>-</td>
            <TC>{fmt(tot.pov_stari/1.2)}</TC><TC>{fmt(tot.pov_novi/1.2)}</TC>
            <TC>{fmt(agg.reduce((s,r)=>s+r.nj_zarada_calc,0))}</TC>
          </tr>
        </tbody>
      </table>
    </div>
    {(()=>{const dugovi=META.sistemi.map(s=>{const d=calcDug(s);return{sistem:s,...d}});return(
      <div style={{marginTop:24}}>
        <SecH>Dugovanja po sistemima</SecH>
        <div style={{overflowX:'auto',border:'1px solid rgba(168,85,247,0.2)',borderTop:'none',borderRadius:'0 0 8px 8px'}}>
          <table><thead><tr><TH>Sistem</TH><TH>Placanje</TH><TH style={{color:'#ec4899'}}>Dug (Faktura)</TH><TH style={{color:'#ec4899'}}>Dug (Odjava)</TH><TH>PDF</TH></tr></thead>
          <tbody>{dugovi.map(d=><tr key={d.sistem} onMouseEnter={e=>e.currentTarget.style.background='rgba(168,85,247,0.05)'} onMouseLeave={e=>e.currentTarget.style.background='transparent'}>
            <TDL>{d.sistem}</TDL>
            <TD style={{textAlign:'center',fontSize:11,color:d.np===1?'#a855f7':'#ec4899'}}>{d.np===1?'Faktura':'Odjava'}</TD>
            <DugCell v={d.dugF}/>
            {d.dugO!==null?<DugCell v={d.dugO}/>:<TD style={{textAlign:'center',color:'#475569',fontSize:11}}>-</TD>}
            <td style={{padding:'7px 12px',borderBottom:'1px solid rgba(168,85,247,0.08)',textAlign:'center'}}>
              {d.np===0&&<button onClick={()=>generateDebtPDF(d.sistem,d.dugF)} style={{padding:'5px 14px',background:'rgba(168,85,247,0.1)',color:'#a855f7',border:'1px solid rgba(168,85,247,0.3)',borderRadius:6,fontSize:11,fontWeight:600,cursor:'pointer',fontFamily:'inherit'}} onMouseEnter={e=>{e.target.style.background='#a855f7';e.target.style.color='#fff'}} onMouseLeave={e=>{e.target.style.background='rgba(168,85,247,0.1)';e.target.style.color='#a855f7'}}>PDF Duga</button>}
            </td>
          </tr>)}</tbody></table>
        </div>
      </div>)})()}
  </div>);
}
function PregledPoSistemu(){
  const allYears=META.godine.filter(y=>y>0);
  const [sistem,setSistem]=useState(META.sistemi[0]);
  const [selYears2,setSelYears2]=useState(allYears.map(String));
  const toggleYear2=y=>{const ys=String(y);const idx=allYears.indexOf(y);setSelYears2(prev=>{if(prev.includes(ys)){return allYears.filter((_,i)=>i<idx).map(String).filter(x=>prev.includes(x))}else{const need=allYears.slice(0,idx+1).map(String);return[...new Set([...prev,...need])]}})};
  const ySet2=useMemo(()=>new Set(selYears2.map(Number)),[selYears2]);
  const monthly=useMemo(()=>DATA.filter(r=>r.sistem===sistem&&ySet2.has(r.godina)).sort((a,b)=>(a.godina*100+a.mesec)-(b.godina*100+b.mesec)),[sistem,ySet2]);
  const pocS=useMemo(()=>RAW.data.filter(r=>r.sistem===sistem&&r.mesec===0).reduce((s,r)=>s+r.poc_stanje,0),[sistem]);
  const tS=k=>monthly.reduce((s,r)=>s+(r[k]||0),0);
  const lNZ=k=>{for(let i=monthly.length-1;i>=0;i--)if(monthly[i][k]!==0)return monthly[i][k];return 0};
  const dug=useMemo(()=>{
    const vProm=monthly.reduce((a,r)=>a+r.v_promet,0)*1.2;
    const pov=monthly.reduce((a,r)=>a+r.pov_total,0);
    const tros=monthly.reduce((a,r)=>a+r.knjizno,0)*1.2+monthly.reduce((a,r)=>a+(r.fakturisano||0),0);
    const upl=monthly.reduce((a,r)=>a+r.uplata,0);
    const uplN=monthly.reduce((a,r)=>a+r.uplata_ka_njima,0);
    const dugF=(pocS+vProm)-(pov+tros+upl+uplN);
    const np=META.nacin_placanja[sistem]||0;
    let dugO=null;
    if(np===0){let lVP=0;for(let i=monthly.length-1;i>=0;i--)if(monthly[i].v_lager_vp!==0){lVP=monthly[i].v_lager_vp;break};dugO=dugF-lVP*1.2}
    return{dugF,dugO,np};
  },[monthly,pocS,sistem]);
  const isO=dug.np===0;
  return(<div style={{padding:'20px'}}>
    <div style={{display:'flex',gap:16,alignItems:'flex-end',marginBottom:20,flexWrap:'wrap'}}>
      <div style={{display:'flex',flexDirection:'column',gap:4}}><Lbl>Sistem</Lbl><select value={sistem} onChange={e=>setSistem(e.target.value)} style={{width:180}}>{META.sistemi.map(s=><option key={s} value={s}>{s}</option>)}</select></div>
      <div style={{display:'flex',flexDirection:'column',gap:4}}><Lbl>Godine</Lbl><div className="ychk">{allYears.map(y=><label key={y}><input type="checkbox" checked={selYears2.includes(String(y))} onChange={()=>toggleYear2(y)}/>{y}</label>)}</div></div>
    </div>
    <div style={{overflowX:'auto',borderRadius:8,border:'1px solid rgba(168,85,247,0.2)',marginBottom:24}}>
      <table style={{minWidth:1200}}>
        <thead>
          <tr><TG colSpan={2}>&nbsp;</TG><TG colSpan={2}>Promet (bez PDV)</TG><TG colSpan={HAS_DOD?4:3}>Trosak (bez PDV)</TG><TG colSpan={2}>Profit</TG><TG colSpan={2}>Vrednost Lagera</TG><TG colSpan={2}>Povrat</TG><TG colSpan={1}>Zarada</TG></tr>
          <tr><TH>Mesec</TH><TH>Godina</TH><TH>Ka sistemu</TH><TH>Ka kupcu</TH><TH>Marketing</TH><TH>Knjizno</TH>{HAS_DOD&&<TH>Dodatni</TH>}<TH>Total</TH><TH>Po prodaji</TH><TH>Po prometu</TH><TH>VP cena</TH><TH>NC cena</TH><TH>Pov. starih</TH><TH>Pov. novih</TH><TH>Njihova</TH></tr>
        </thead>
        <tbody>
          {monthly.map((r,i)=><tr key={r.godina+'-'+r.mesec} onMouseEnter={e=>e.currentTarget.style.background='rgba(168,85,247,0.05)'} onMouseLeave={e=>e.currentTarget.style.background='transparent'}>
            <TDL>{ML[r.mesec]||r.mesec}</TDL><TD style={{color:'#8a919e',textAlign:'center',fontSize:12}}>{r.godina}</TD>
            <TD>{fmt(r.v_promet)}</TD><TD>{fmt(r.v_kupac)}</TD><TD>{fmt(r.marketing)}</TD><TD>{fmt(r.knjizno)}</TD>{HAS_DOD&&<TD>{fmt(r.dodatni_trosak||0)}</TD>}<TD>{fmt(r.total_trosak)}</TD><TD neg>{fmt(r.profit_prodaja)}</TD><TD neg>{fmt(r.profit_promet)}</TD><TD>{fmt(r.v_lager_vp)}</TD><TD>{fmt(r.v_lager_nc)}</TD><TD>{fmt(r.pov_stari/1.2)}</TD><TD>{fmt(r.pov_novi/1.2)}</TD><TD>{fmt(calcNjZarada(monthly,i))}</TD>
          </tr>)}
          <tr style={{background:'rgba(168,85,247,0.08)'}}>
            <td colSpan={2} style={{padding:'9px 12px',fontWeight:700,color:'#a855f7',fontSize:13,position:'sticky',left:0,background:'rgba(168,85,247,0.08)'}}>UKUPNO</td>
            {['v_promet','v_kupac','marketing','knjizno'].map(k=><TC key={k}>{fmt(tS(k))}</TC>)}
            {HAS_DOD&&<TC>{fmt(tS('dodatni_trosak'))}</TC>}
            <TC>{fmt(tS('total_trosak'))}</TC>
            {['profit_prodaja','profit_promet'].map(k=>{const v=tS(k);return<TC key={k} color={v>=0?'#a855f7':'#ec4899'}>{fmt(v)}</TC>})}
            {['v_lager_vp','v_lager_nc'].map(k=><TC key={k}>{fmt(lNZ(k))}</TC>)}
            {['pov_stari','pov_novi'].map(k=><TC key={k}>{fmt(tS(k)/1.2)}</TC>)}
            <TC>{fmt(calcNjZaradaTotal(monthly))}</TC>
          </tr>
        </tbody>
      </table>
    </div>
    <div style={{marginBottom:24}}>
      <SecH extra={isO&&<button onClick={()=>generateDebtPDF(sistem,dug.dugF)} style={{padding:'6px 18px',background:'rgba(168,85,247,0.1)',color:'#a855f7',border:'1px solid rgba(168,85,247,0.3)',borderRadius:6,fontSize:12,fontWeight:600,cursor:'pointer',fontFamily:'inherit'}} onMouseEnter={e=>{e.target.style.background='#a855f7';e.target.style.color='#fff'}} onMouseLeave={e=>{e.target.style.background='rgba(168,85,247,0.1)';e.target.style.color='#a855f7'}}>&#128196; Generisi PDF duga</button>}>Finansijski Bilans (sa PDV)</SecH>
      <div style={{overflowX:'auto',border:'1px solid rgba(168,85,247,0.2)',borderTop:'none',borderRadius:'0 0 8px 8px'}}>
        <table style={{minWidth:900}}><thead><tr>
          <TH>Poc. stanje</TH><TH>Promet+PDV</TH><TH style={{color:'#a855f7'}}>ZADUZENJE</TH><TH>Povrat(PDV)</TH><TH>Trosak+PDV</TH><TH style={{color:'#a855f7'}}>SA UMANJENJEM</TH><TH>Uplata</TH><TH>Uplata ka njima</TH><TH style={{color:'#ec4899'}}>Dug (Faktura)</TH>
          {dug.dugO!==null&&<React.Fragment><TH>Lager VP(PDV)</TH><TH style={{color:'#ec4899'}}>Dug (Odjava)</TH></React.Fragment>}
        </tr></thead>
        <tbody><tr>
          <TD>{fmt(pocS)}</TD><TD>{fmt(tS('v_promet')*1.2)}</TD><TD style={{color:'#a855f7',fontWeight:700}}>{fmt(pocS+tS('v_promet')*1.2)}</TD><TD>{fmt(tS('pov_total'))}</TD><TD>{fmt(tS('knjizno')*1.2+tS('fakturisano'))}</TD><TD style={{color:'#a855f7',fontWeight:700}}>{fmt(pocS+tS('v_promet')*1.2-tS('pov_total')-(tS('knjizno')*1.2+tS('fakturisano')))}</TD><TD>{fmt(tS('uplata'))}</TD><TD>{fmt(tS('uplata_ka_njima'))}</TD><DugCell v={dug.dugF}/>
          {dug.dugO!==null&&<React.Fragment><TD>{fmt(lNZ('v_lager_vp')*1.2)}</TD><DugCell v={dug.dugF-lNZ('v_lager_vp')*1.2}/></React.Fragment>}
        </tr></tbody></table>
      </div>
    </div>
  </div>);
}
function Dashboard(){
  const [tab,setTab]=useState('total');
  return(<div style={{minHeight:'100vh',background:'#12002a',color:'#c9d1d9',fontFamily:"'DM Sans',-apple-system,sans-serif"}}>
    <div style={{maxWidth:1600,margin:'0 auto',padding:'20px'}}>
      <div style={{display:'flex',borderBottom:'1px solid rgba(168,85,247,0.2)',marginBottom:20}}>
        <TabBtn active={tab==='total'} onClick={()=>setTab('total')}>Total Pregled</TabBtn>
        <TabBtn active={tab==='sistem'} onClick={()=>setTab('sistem')}>Pregled po Sistemu</TabBtn>
      </div>
      {tab==='total'?<TotalPregled/>:<PregledPoSistemu/>}
    </div>
  </div>);
}
ReactDOM.createRoot(document.getElementById('root')).render(<Dashboard/>);
</script>
</body>
</html>'''
        return html_out

    with st.spinner("⏳ Učitavam podatke..."):
        html_content = build_finansijski_html()

    if html_content is None:
        st.error("❌ Podaci nisu dostupni. Proveri da li je sistemi.xlsx postavljen na GitHub.")
    else:
        components.html(html_content, height=900, scrolling=True)
