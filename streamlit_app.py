import streamlit as st
import streamlit.components.v1 as components
import io, datetime, math, numpy as np, pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter

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
        self.log(f"Prodaja: {len(self.prodaja)} redova")
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
            if lager<=2 and pred>0:
                target=max(pred, prosek, self.min_order)
                dopuna=max(target-lager,0)
            elif lager<=2 and pred==0: dopuna=0
            else: dopuna=max(self.min_lager-lager,0)
            return max(osnova,dopuna)
        self.df_result['Porudzbina_1']=self.df_result.apply(p1,axis=1).astype(int)
        self.df_result['Porudzbina_2']=self.df_result.apply(p2,axis=1).astype(int)

        # --- FINALNA PROVERA: porudzbina_2 + lager ne sme biti <= prodaja poslednjeg meseca ---
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
        self.adjustments=[]

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

        # Zastita: ako nema poklapanja, koristi sve mesece
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

st.set_page_config(page_title="VAPE Analitika", page_icon="\U0001f4a8", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap');
    .stApp { background: linear-gradient(160deg, #fdf2f8 0%, #f5f0ff 40%, #eff6ff 100%); font-family: 'Poppins', sans-serif; }
    section[data-testid="stSidebar"] { background: linear-gradient(180deg, #7c3aed 0%, #a855f7 50%, #c084fc 100%) !important; }
    section[data-testid="stSidebar"] * { color: white !important; }
    section[data-testid="stSidebar"] input, section[data-testid="stSidebar"] textarea {
        background: rgba(255,255,255,0.9) !important; border: 1px solid rgba(255,255,255,0.3) !important;
        color: #1a1a2e !important; border-radius: 8px !important; }
    .metric-card { background: white; border-radius: 16px; padding: 16px 20px;
        box-shadow: 0 2px 12px rgba(124,58,237,0.08); border: 1px solid rgba(124,58,237,0.1); text-align: center; }
    .metric-value { font-size: 26px; font-weight: 700;
        background: linear-gradient(135deg, #7c3aed, #ec4899); -webkit-background-clip: text; -webkit-text-fill-color: transparent; }
    .metric-value-red { font-size: 26px; font-weight: 700; color: #dc2626; }
    .metric-value-green { font-size: 26px; font-weight: 700; color: #059669; }
    .metric-label { font-size: 11px; color: #888; margin-top: 4px; }
    .stButton > button { background: linear-gradient(135deg, #a855f7 0%, #ec4899 100%) !important;
        color: white !important; border: none !important; border-radius: 12px !important;
        padding: 12px 32px !important; font-weight: 600 !important; font-size: 16px !important;
        box-shadow: 0 4px 15px rgba(168,85,247,0.3) !important; }
    [data-testid="stDownloadButton"]:first-of-type > button { background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important;
        color: white !important; border: none !important; border-radius: 12px !important;
        padding: 12px 32px !important; font-weight: 600 !important;
        box-shadow: 0 4px 15px rgba(16,185,129,0.3) !important; }
    .stDownloadButton > button { background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important;
        color: white !important; border: none !important; border-radius: 12px !important;
        padding: 12px 32px !important; font-weight: 600 !important;
        box-shadow: 0 4px 15px rgba(16,185,129,0.3) !important; }
    .header-banner { background: linear-gradient(135deg, #7c3aed 0%, #a855f7 30%, #ec4899 70%, #f472b6 100%);
        border-radius: 16px; padding: 24px 32px; color: white; margin-bottom: 24px;
        box-shadow: 0 4px 20px rgba(124,58,237,0.25); }
    .header-title { font-size: 28px; font-weight: 700; margin: 0; }
    .header-sub { font-size: 14px; opacity: 0.85; margin-top: 4px; }
    .success-box { background: linear-gradient(135deg, rgba(16,185,129,0.1), rgba(5,150,105,0.05));
        border: 1px solid rgba(16,185,129,0.2); border-radius: 12px; padding: 16px 20px; }
    .warn-box { background: linear-gradient(135deg, rgba(220,38,38,0.08), rgba(220,38,38,0.03));
        border: 1px solid rgba(220,38,38,0.2); border-radius: 12px; padding: 12px 16px; margin: 8px 0; }
    .section-title { font-size: 18px; font-weight: 600; color: #4c1d95; margin: 16px 0 8px 0; }
</style>
""", unsafe_allow_html=True)

st.markdown("""<div class="header-banner"><div class="header-title">\U0001f4a8 VAPE ANALITIKA & PORUDZBINE</div>
    <div class="header-sub">Predikcija prodaje \u2022 Profitabilnost \u2022 OOS analiza \u2022 Efekti akcije</div></div>""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### \u2699\ufe0f Parametri modela")
    alpha = st.number_input("Alpha (nivo)", 0.0, 1.0, 0.4, 0.05)
    beta = st.number_input("Beta (trend)", 0.0, 1.0, 0.2, 0.05)
    min_lager = st.number_input("Min lager", 0, 20, 2)
    min_order = st.number_input("Min porudzbina po objektu", 0, 50, 2)
    st.markdown("---")
    st.markdown("### \U0001f4b0 Troskovi")
    mesecni_trosak = st.number_input(
        "Ukupan trosak mkt/ulistavanja za ceo period (RSD)",
        min_value=0, value=0, step=10000,
        help="Unesi UKUPAN iznos za ceo analizirani period — automatski se deli na broj objekata i broj meseci"
    )
    st.markdown("---")
    st.markdown("### \u26d4 Iskljuceni komitenti")
    excluded_str = st.text_area("ID-evi razdvojeni zarezom", value=DEFAULT_EXCLUDED, height=100)

excluded = set()
for part in excluded_str.replace('\n', ',').split(','):
    p = part.strip()
    if p.isdigit(): excluded.add(int(p))

uploaded = st.file_uploader("Ucitaj Excel fajl sa podacima", type=['xlsx','xls'])

if uploaded:
    file_bytes = uploaded.read()
    st.markdown(f'<div class="success-box">\u2705 Fajl <strong>{uploaded.name}</strong> ucitan ({len(file_bytes)//1024} KB)</div>', unsafe_allow_html=True)
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
            _labels = [f"{_mn.get(int(m),'?')} {int(g)}" for g,m in _meseci]
            st.markdown("**\U0001f4c5 Period za analizu** (OOS, Profitabilnost, Akcija — ne utice na predikciju):")
            selected_labels = st.multiselect("Odaberi mesece", _labels, default=_labels, help="Predikcija uvek koristi sve mesece. Ovaj filter se odnosi samo na analitiku.")
            if not selected_labels:
                st.warning("⚠️ Mora biti odabran bar jedan mesec za analizu. Automatski je odabran poslednji mesec.")
                selected_labels = [_labels[-1]] if _labels else []
            selected_meseci = [_meseci[i] for i, lb in enumerate(_labels) if lb in selected_labels]
        else:
            selected_labels = []; selected_meseci = []
    except:
        selected_labels = []; selected_meseci = []

    if st.button("\U0001f680 POKRENI ANALIZU", use_container_width=True):
        progress_bar = st.progress(0)
        try:
            engine = PredictionEngine(file_bytes, excluded, alpha, beta, min_lager, min_order, mesecni_trosak, selected_meseci)
            result = engine.run(progress_bar)

            st.markdown("---")
            tp = int(result['Predikcija'].sum()); tl = int(result['Lager_danas'].sum())
            t1 = int(result[~result['ID KOMITENTA'].isin(excluded)]['Porudzbina_1'].sum())
            t2 = int(result[~result['ID KOMITENTA'].isin(excluded)]['Porudzbina_2'].sum())

            if engine.has_prices:
                total_profit = int(engine.df_profit_obj['Bruto_profit'].sum())
                total_lost = int(engine.df_oos['Izgubljeni_profit'].sum()) if len(engine.df_oos) > 0 else 0
                n_unprofitable = int((engine.df_profit_obj['Neto_profit'] <= 0).sum())
                c1,c2,c3,c4,c5,c6 = st.columns(6)
                c1.markdown(f'<div class="metric-card"><div class="metric-value">{tp:,}</div><div class="metric-label">Predikcija (kom)</div></div>', unsafe_allow_html=True)
                c2.markdown(f'<div class="metric-card"><div class="metric-value">{t2:,}</div><div class="metric-label">Porudzbina P2</div></div>', unsafe_allow_html=True)
                c3.markdown(f'<div class="metric-card"><div class="metric-value">{tl:,}</div><div class="metric-label">Lager</div></div>', unsafe_allow_html=True)
                c4.markdown(f'<div class="metric-card"><div class="metric-value-green">{total_profit:,}</div><div class="metric-label">Bruto profit (RSD)</div></div>', unsafe_allow_html=True)
                c5.markdown(f'<div class="metric-card"><div class="metric-value-red">-{total_lost:,}</div><div class="metric-label">Izgubljeno OOS (RSD)</div></div>', unsafe_allow_html=True)
                c6.markdown(f'<div class="metric-card"><div class="metric-value-red">{n_unprofitable}</div><div class="metric-label">Neprofitabilnih obj.</div></div>', unsafe_allow_html=True)
            else:
                c1,c2,c3,c4,c5 = st.columns(5)
                for col, val, lbl in [(c1,tp,'Predikcija'),(c2,tl,'Lager'),(c3,t1,'P1'),(c4,t2,f'P2 min {min_lager}'),(c5,result[result['Porudzbina_2']>0]['ID KOMITENTA'].nunique(),'Objekata')]:
                    col.markdown(f'<div class="metric-card"><div class="metric-value">{val:,}</div><div class="metric-label">{lbl}</div></div>', unsafe_allow_html=True)

            st.markdown("")

            if engine.has_prices:
                tab1, tab2, tab4, tab5 = st.tabs(["\U0001f4e6 Porudzbina", "\U0001f4b0 Profitabilnost objekata & OOS", "\U0001f3af Analiza Akcije", "\U0001f4cb Log"])
            else:
                tab1, tab5 = st.tabs(["\U0001f4e6 Porudzbina", "\U0001f4cb Log"])

            with tab1:
                cols_show = ['ID KOMITENTA','id artikla','Naziv artikla','Grupa']
                if engine.has_history: cols_show.append('Total_JanAvg')
                cols_show += ['Predikcija','Prosek','Lager_danas','Porudzbina_1','Porudzbina_2']
                show = result[cols_show].copy()
                names = ['ID Kom.','ID Art.','Naziv','Grupa']
                if engine.has_history: names.append('Jan-Avg')
                names += ['Pred.','Prosek','Lager','P1','P2']
                show.columns = names
                st.dataframe(show, use_container_width=True, height=400)

            if engine.has_prices:
                with tab2:
                    period_str2 = ", ".join(engine.analitika_labels) if engine.analitika_labels else "svi meseci"
                    n_mes = len(engine.analitika_labels) if engine.analitika_labels else len(engine.mesec_labels)
                    n_obj = engine.num_komitenti

                    # --- Izracunaj ukupne vrednosti za summary kartice ---
                    prof = engine.df_profit_obj.copy()
                    total_bruto = int(prof['Bruto_profit'].sum())
                    total_neto = int(prof['Neto_profit'].sum())
                    total_trosak = int(prof['Trosak_mkt'].sum())
                    total_oos_izgubljen = int(engine.df_oos['Izgubljeni_profit'].sum()) if len(engine.df_oos) > 0 else 0
                    mes_trosak = total_trosak / max(n_mes, 1)
                    mes_bruto = total_bruto / max(n_mes, 1)
                    mes_neto = total_neto / max(n_mes, 1)
                    mes_oos = total_oos_izgubljen / max(n_mes, 1)

                    st.caption(f"\U0001f4c5 Period analize: **{period_str2}** · {n_obj} objekata · {n_mes} meseci")

                    # --- 4 summary kartice ---
                    ka, kb, kc, kd = st.columns(4)
                    def _kard(col, label, total, mes, color, prefix=""):
                        col.markdown(f"""
                        <div style="background:white;border-radius:12px;padding:16px 18px;
                            border-left:4px solid {color};box-shadow:0 2px 8px rgba(0,0,0,0.07);height:100%;">
                            <div style="font-size:10px;color:#999;font-weight:600;letter-spacing:.5px;text-transform:uppercase;margin-bottom:6px;">{label}</div>
                            <div style="font-size:22px;font-weight:700;color:{color};">{prefix}{total:,.0f} RSD</div>
                            <div style="font-size:11px;color:#aaa;margin-top:3px;">{prefix}{mes:,.0f} RSD / mesec</div>
                        </div>""", unsafe_allow_html=True)
                    _kard(ka, f"Ukupan trosak · {n_mes} meseci", total_trosak, mes_trosak, "#f59e0b")
                    _kard(kb, f"Bruto profit · {n_mes} meseci", total_bruto, mes_bruto, "#10b981")
                    _kard(kc, f"Neto profit · {n_mes} meseci", total_neto, mes_neto, "#3b82f6" if total_neto > 0 else "#ef4444")
                    _kard(kd, f"OOS izgubljen · {n_mes} meseci", total_oos_izgubljen, mes_oos, "#ef4444", prefix="-")

                    st.markdown("<div style='margin:20px 0 4px 0;'></div>", unsafe_allow_html=True)

                    # --- Mesecni trendovi: bruto i neto profit ---
                    # Sakupi mesecne podatke iz df_profit_obj
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
                        boja = "#10b981" if promena_pct > 0 else "#ef4444"
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
                                <div style="flex:1;background:#f3f4f6;border-radius:3px;height:18px;position:relative;">
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
                        chart_b = _bar_chart_html(bruto_po_mes, max_bruto, "#10b981", "#ef4444")
                        components.html(f'<!DOCTYPE html><html><body style="margin:0;padding:8px 12px;font-family:sans-serif;">{chart_b}</body></html>', height=len(bruto_po_mes)*28+20)

                    with col_neto:
                        st.markdown('<div class="section-title">📉 Mesečni trend neto profita</div>', unsafe_allow_html=True)
                        rec_n = _trend_recenica(neto_po_mes, "Neto profit")
                        if rec_n: st.markdown(f'<p style="font-size:13px;color:#555;margin-bottom:6px;">{rec_n}</p>', unsafe_allow_html=True)
                        chart_n = _bar_chart_html(neto_po_mes, max_neto, "#3b82f6", "#ef4444")
                        components.html(f'<!DOCTYPE html><html><body style="margin:0;padding:8px 12px;font-family:sans-serif;">{chart_n}</body></html>', height=len(neto_po_mes)*28+20)

                    st.markdown("<div style='margin:20px 0 4px 0;'></div>", unsafe_allow_html=True)

                    # --- Profitabilnost po objektima ---
                    st.markdown('<div class="section-title">🏪 Profitabilnost po objektima</div>', unsafe_allow_html=True)

                    # Izracunaj kljucne brojeve
                    ukupno_obj = len(prof)
                    neto_neg = prof[prof['Neto_profit'] <= 0]
                    n_neto_neg = len(neto_neg)
                    # OOS-negativni: neto_profit <= 0 ali potencijal > 0
                    oos_neg = prof[(prof['Neto_profit'] <= 0) & (prof['Potencijalni_profit'] > 0)]
                    n_oos_neg = len(oos_neg)
                    # Pravi neprofitabilni: negativni i po potencijalu
                    pravi_neg = prof[(prof['Neto_profit'] <= 0) & (prof['Potencijalni_profit'] <= 0)]
                    n_pravi_neg = len(pravi_neg)
                    pct_pravi = round(n_pravi_neg / max(ukupno_obj, 1) * 100)
                    trosak_po_obj = engine.trosak_po_objektu
                    trosak_mes_obj = trosak_po_obj / max(n_mes, 1)
                    usteda_trosak = n_pravi_neg * trosak_po_obj
                    usteda_gubitak = abs(pravi_neg['Neto_profit'].sum()) if n_pravi_neg > 0 else 0
                    usteda_mes = (usteda_trosak + usteda_gubitak) / max(n_mes, 1)

                    # Analiticki tekst
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
                    st.markdown(tekst, unsafe_allow_html=True)

                    # --- Grafikon: profitabilni vs neprofitabilni po mesecima ---
                    a_labels_trend2 = engine.analitika_labels if engine.analitika_labels else engine.mesec_labels
                    a_meseci_trend2 = engine.analitika_meseci if (engine.analitika_meseci and len(engine.analitika_meseci) > 0) else engine.meseci_order

                    chart_mes_data = []
                    for i, (lb, (g, m)) in enumerate(zip(a_labels_trend2, a_meseci_trend2)):
                        col_neto_lb = f'Neto_{lb}'
                        if col_neto_lb in prof.columns:
                            n_prof_mes = (prof[col_neto_lb] > 0).sum()
                            n_nepr_mes = (prof[col_neto_lb] <= 0).sum()
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
                                    <div style="width:{bar_w}px;height:{h_p}px;background:#10b981;border-radius:3px 3px 0 0;position:relative;">
                                        <span style="position:absolute;top:-18px;left:50%;transform:translateX(-50%);font-size:10px;font-weight:700;color:#10b981;white-space:nowrap;">{np_v}</span>
                                    </div>
                                    <div style="width:{bar_w}px;height:{h_n}px;background:#ef4444;border-radius:3px 3px 0 0;position:relative;">
                                        <span style="position:absolute;top:-18px;left:50%;transform:translateX(-50%);font-size:10px;font-weight:700;color:#ef4444;white-space:nowrap;">{nn_v}</span>
                                    </div>
                                </div>
                                <div style="font-size:10px;color:#888;margin-top:4px;text-align:center;width:{bar_w*2+3}px;">{lb}</div>
                            </div>"""
                        chart_html = f"""<!DOCTYPE html><html><body style="margin:0;padding:0;font-family:sans-serif;background:white;">
                        <div style="padding:16px 20px;">
                            <div style="display:flex;gap:16px;margin-bottom:14px;">
                                <span style="display:flex;align-items:center;gap:5px;font-size:12px;color:#555;">
                                    <span style="width:12px;height:12px;background:#10b981;border-radius:2px;display:inline-block;"></span> Profitabilni
                                </span>
                                <span style="display:flex;align-items:center;gap:5px;font-size:12px;color:#555;">
                                    <span style="width:12px;height:12px;background:#ef4444;border-radius:2px;display:inline-block;"></span> Neprofitabilni
                                </span>
                            </div>
                            <div style="display:flex;gap:6px;align-items:flex-end;overflow-x:auto;padding-bottom:4px;">
                                {bars_html}
                            </div>
                        </div>
                        </body></html>"""
                        components.html(chart_html, height=220)

                    unprofitable = prof[prof['Neto_profit'] <= 0].sort_values('Neto_profit')
                    if len(unprofitable) > 0:
                        with st.expander(f"⚠️ {len(unprofitable)} neprofitabilnih objekata — kandidati za izlistavanje"):
                            up_show = unprofitable[['ID KOMITENTA','Artikala','Prodato_kom','Bruto_profit','Trosak_mkt','Neto_profit','Izgubljeno_OOS','Potencijalni_profit']].copy()
                            up_show.columns = ['ID Kom.','Art.','Prod. kom','Bruto profit','Trosak mkt','Neto profit','Izg. OOS','Potencijal']
                            st.dataframe(up_show, use_container_width=True, height=250)

                    # Download profitabilnih objekata
                    profitabilni = prof[prof['Neto_profit'] > 0].sort_values('Neto_profit', ascending=False)
                    if len(profitabilni) > 0:
                        import io as _io
                        prof_buf = _io.BytesIO()
                        prof_show = profitabilni[['ID KOMITENTA','Artikala','Prodato_kom','Bruto_profit','Trosak_mkt','Neto_profit','Izgubljeno_OOS','Potencijalni_profit']].copy()
                        prof_show.columns = ['ID Kom.','Art.','Prod. kom','Bruto profit','Trosak mkt','Neto profit','Izg. OOS','Potencijal']
                        with pd.ExcelWriter(prof_buf, engine='openpyxl') as _w:
                            prof_show.to_excel(_w, index=False, sheet_name='Profitabilni objekti')
                        prof_buf.seek(0)
                        st.download_button(f"⬇️ Preuzmi listu profitabilnih objekata ({len(profitabilni)})",
                            data=prof_buf, file_name=f"profitabilni_objekti_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                    st.markdown("<div style='margin:20px 0 4px 0;'></div>", unsafe_allow_html=True)

                    # --- OOS sekcija ---
                    st.markdown('<div class="section-title">🔴 OOS — Izgubljena zarada zbog nedostatka robe</div>', unsafe_allow_html=True)
                    if len(engine.df_oos) > 0:
                        a_labels_oos = engine.analitika_labels if engine.analitika_labels else engine.mesec_labels

                        # --- 3 summary kartice ---
                        oos_ukupno = int(engine.df_oos['Izgubljeni_profit'].sum())
                        oos_mes_avg = oos_ukupno // max(n_mes, 1)
                        oos_kombinacija = int((engine.df_oos['OOS_meseci'] > 0).sum()) if 'OOS_meseci' in engine.df_oos.columns else len(engine.df_oos)
                        # broji kombinacije na 0 lagera danas
                        oos_0_danas = int((engine.df_oos.get('Lager_danas', 0) == 0).sum()) if 'Lager_danas' in engine.df_oos.columns else oos_kombinacija

                        o1, o2, o3 = st.columns(3)
                        def _oos_kard(col, label, val, suffix=""):
                            col.markdown(f"""<div style="background:white;border-radius:12px;padding:16px 18px;
                                border-top:3px solid #ef4444;box-shadow:0 2px 8px rgba(0,0,0,0.07);text-align:center;">
                                <div style="font-size:22px;font-weight:700;color:#ef4444;">{val:,}{suffix}</div>
                                <div style="font-size:11px;color:#aaa;margin-top:4px;text-transform:uppercase;letter-spacing:.5px;">{label}</div>
                            </div>""", unsafe_allow_html=True)
                        _oos_kard(o1, f"Izgubljen profit · {n_mes} meseci (RSD)", oos_ukupno)
                        _oos_kard(o2, "Prosečno mesečno (RSD)", oos_mes_avg)
                        _oos_kard(o3, "Kombinacija na 0 lagera danas", oos_0_danas)

                        st.markdown("<div style='margin:18px 0 4px 0;'></div>", unsafe_allow_html=True)

                        # --- Line chart: izgubljen profit + broj OOS kombinacija po mesecu ---
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

                            # area fill
                            pts_area = " ".join(f"{px(i)},{py(v)}" for i, v in enumerate(mes_izgub))
                            pts_area = f"{px(0)},{pad_t+plot_h} " + pts_area + f" {px(n_pts-1)},{pad_t+plot_h}"
                            # line
                            pts_line = " ".join(f"{px(i)},{py(v)}" for i, v in enumerate(mes_izgub))

                            dots = ""
                            labels_svg = ""
                            x_labels = ""
                            for i, (lb, v, vc) in enumerate(zip(a_labels_oos, mes_izgub, mes_oos_count)):
                                x, y = px(i), py(v)
                                v_k = f"{v//1000}k" if v >= 1000 else str(v)
                                dots += f'<circle cx="{x}" cy="{y}" r="5" fill="#ef4444" stroke="white" stroke-width="2"/>'
                                labels_svg += f'<text x="{x}" y="{y-10}" text-anchor="middle" font-size="10" font-weight="700" fill="#ef4444">{v_k}</text>'
                                labels_svg += f'<text x="{x}" y="{y+20}" text-anchor="middle" font-size="9" fill="#999">({vc})</text>'
                                x_labels += f'<text x="{x}" y="{chart_h-6}" text-anchor="middle" font-size="9" fill="#aaa">{lb}</text>'

                            svg = f"""<svg width="{chart_w}" height="{chart_h}" xmlns="http://www.w3.org/2000/svg" style="font-family:sans-serif;">
                                <text x="{pad_l-5}" y="{pad_t-8}" font-size="10" fill="#888">Izgubljen profit (RSD)</text>
                                <text x="{chart_w-pad_r}" y="{pad_t-8}" font-size="10" fill="#aaa" text-anchor="end">Broj OOS kombinacija u zagradama</text>
                                <polygon points="{pts_area}" fill="#ef4444" fill-opacity="0.08"/>
                                <polyline points="{pts_line}" fill="none" stroke="#ef4444" stroke-width="2.5"/>
                                {dots}{labels_svg}{x_labels}
                            </svg>"""
                            components.html(f'<!DOCTYPE html><html><body style="margin:0;padding:0;background:white;">{svg}</body></html>', height=chart_h+10)

                        # --- Top 5 artikala po izgubljenom profitu ---
                        oos_art = engine.df_oos.groupby(['id artikla','Naziv artikla']).agg(
                            Izgubljeni_profit=('Izgubljeni_profit','sum')
                        ).reset_index().sort_values('Izgubljeni_profit', ascending=False).head(5)

                        bar_colors = ["#ef4444","#f97316","#ec4899","#8b5cf6","#3b82f6"]
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
                                <div style="flex:1;background:#f3f4f6;border-radius:4px;height:22px;position:relative;">
                                    <div style="width:{pct:.1f}%;background:{color};height:100%;border-radius:4px;"></div>
                                </div>
                                <div style="width:110px;font-size:12px;font-weight:700;color:{color};flex-shrink:0;">{val:,} RSD</div>
                            </div>"""

                        st.markdown("**Top 5 artikala po izgubljenom profitu:**")
                        components.html(f"""<!DOCTYPE html><html><body style="margin:0;padding:8px 12px;font-family:sans-serif;background:white;">
                            {bars5}
                        </body></html>""", height=len(oos_art)*42+20)

                        # --- Samo artikli expander ---
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

                    # --- SCENARIO: Optimalna mreza ---
                    st.markdown("<div style='margin:24px 0 4px 0;'></div>", unsafe_allow_html=True)
                    st.markdown('<div class="section-title">⚡ Scenario: Optimalna mreža</div>', unsafe_allow_html=True)

                    # Izracunaj za odabrani period
                    prof2 = engine.df_profit_obj.copy()
                    oos_ukupno2 = int(engine.df_oos['Izgubljeni_profit'].sum()) if len(engine.df_oos) > 0 else 0

                    # Pozitivni objekti (potencijal > 0)
                    pozitivni = prof2[prof2['Potencijalni_profit'] > 0]
                    neto_pozitivnih = int(pozitivni['Neto_profit'].sum())

                    # Pravi neprofitabilni (negativni i po potencijalu)
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
                        <div style="font-size:12px;font-weight:600;color:#8b5cf6;margin-bottom:12px;
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
                        {_red(f"Razlika — potencijal koji još nije ostvaren", razlika, "#8b5cf6")}
                    </div>"""
                    st.markdown(scenario_html, unsafe_allow_html=True)

                with tab4:
                    period_str3 = ", ".join(engine.analitika_labels) if engine.analitika_labels else "svi meseci"
                    st.markdown('<div class="section-title">\U0001f3af Efekat akcijske cene & Obrt lagera</div>', unsafe_allow_html=True)
                    st.caption(f"\U0001f4c5 Period analize: **{period_str3}**")
                    if len(engine.df_promo) > 0:
                        promo = engine.df_promo
                        total_akcija = int(promo['Profit_akcija'].sum())
                        total_cena = int(promo['Cena_akcije'].sum())
                        total_prihod_akc = int(promo['Prihod_akcija'].sum())
                        total_prihod_red = int(promo['Prihod_redovna'].sum())
                        avg_obrt = promo['Obrt_x'].mean()
                        cc1, cc2, cc3, cc4 = st.columns(4)
                        cc1.markdown(f'<div class="metric-card"><div class="metric-value-green">{total_prihod_akc:,}</div><div class="metric-label">Prihod na akciji (RSD)</div></div>', unsafe_allow_html=True)
                        cc2.markdown(f'<div class="metric-card"><div class="metric-value-green">{total_akcija:,}</div><div class="metric-label">Profit na akciji (RSD)</div></div>', unsafe_allow_html=True)
                        cc3.markdown(f'<div class="metric-card"><div class="metric-value-red">-{total_cena:,}</div><div class="metric-label">Cena akcije (RSD)</div></div>', unsafe_allow_html=True)
                        cc4.markdown(f'<div class="metric-card"><div class="metric-value">{avg_obrt:.1f}x</div><div class="metric-label">Prosecni obrt lagera</div></div>', unsafe_allow_html=True)
                        st.markdown("")
                        st.markdown("**Pregled po artiklima** (sortirano po obrtu lagera):")
                        pr_show = promo[['id artikla','Naziv','Grupa','Popust_%','Prodato_kom',
                                         'Prihod_akcija','Profit_akcija','Cena_akcije',
                                         'Avg_lager','Obrt_x','Dani_pokrivanja',
                                         'Obj_aktivnih','Prod_po_obj']].copy()
                        pr_show.columns = ['ID Art.','Naziv','Grupa','Popust %','Prod. kom',
                                           'Prihod akcija','Profit akcija','Cena akcije',
                                           'Pros. lager','Obrt (x)','Dani pokr.',
                                           'Akt. obj.','Prod/obj']
                        st.dataframe(pr_show, use_container_width=True, height=350)
                        best = promo.iloc[0]; worst = promo.iloc[-1]
                        st.markdown(f"""
**Uvidi:**
- Najbolji obrt: **{best['Naziv'][:40]}** — {best['Obrt_x']}x obrt, {int(best['Dani_pokrivanja'])} dana pokrivanja, {best['Prod_po_obj']} kom/obj
- Najslabiji obrt: **{worst['Naziv'][:40]}** — {worst['Obrt_x']}x obrt, {int(worst['Dani_pokrivanja'])} dana pokrivanja, {worst['Prod_po_obj']} kom/obj
- Akcija je kostala **{total_cena:,} RSD** u izgubljenom profitu (razlika akcijska vs redovna cena)
- Prihod bi na redovnoj ceni bio **{total_prihod_red:,} RSD** umesto {total_prihod_akc:,} RSD
""")

            with tab5:
                for msg in engine.logs: st.text(msg)

            st.markdown("---")
            excel_buf = create_excel(engine)
            fname_xl = f"ANALITIKA_{datetime.date.today().strftime('%Y%m%d')}.xlsx"
            st.download_button(f"\U0001f4e5 Preuzmi Excel — {fname_xl}", data=excel_buf, file_name=fname_xl,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

        except Exception as e:
            st.error(f"Greska: {str(e)}")
            import traceback; st.code(traceback.format_exc())
else:
    st.markdown("""<div style="text-align:center;padding:60px 20px;color:#aaa;">
        <div style="font-size:48px;margin-bottom:12px;">\U0001f4c2</div>
        <div style="font-size:16px;color:#888;">Ucitaj Excel fajl da pocnes</div>
        <div style="font-size:12px;color:#bbb;margin-top:8px;">Sheetovi: prodaja, startni lager, povrat, trenutni lager, prodaja pre septembra</div>
        <div style="font-size:12px;color:#bbb;">Opciono: kolone Redovna cena, Akcijska cena, Finalna cena, Nabavna vrednost, Profit</div>
    </div>""", unsafe_allow_html=True)
