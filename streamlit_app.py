import streamlit as st
import io, datetime, numpy as np, pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

WMA_WEIGHTS = np.array([0.05, 0.10, 0.15, 0.25, 0.45])
HIST_WEIGHT = 0.05

class PredictionEngine:
    def __init__(self, file_bytes, excluded_ids, alpha, beta, min_lager, min_order):
        self.file_bytes = file_bytes; self.excluded = excluded_ids
        self.alpha = alpha; self.beta = beta; self.min_lager = min_lager; self.min_order = min_order
        self.logs = []; self.adjustments = []; self.has_history = False

    def log(self, msg): self.logs.append(msg)

    def run(self, progress_bar):
        progress_bar.progress(5, "Ucitavanje..."); self._load_sheets()
        progress_bar.progress(15, "Priprema..."); self._prepare_lookups()
        progress_bar.progress(25, "Povrat/korekcija..."); self._compute_povrat()
        progress_bar.progress(40, "Mesecni pregled..."); self._build_monthly()
        progress_bar.progress(55, "Predikcija..."); self._predict_all()
        progress_bar.progress(70, "Lager..."); self._merge_lager()
        progress_bar.progress(80, "Porudzbine..."); self._compute_orders()
        progress_bar.progress(85, "Min pravilo..."); self._apply_min_order()
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
        self.povrat_df = pd.DataFrame()
        if s_pov:
            self.povrat_df = pd.read_excel(xls, sheet_name=s_pov); self.povrat_df.columns=[c.strip() for c in self.povrat_df.columns]
            self.log(f"Povrat: {len(self.povrat_df)} redova")
        self.trenutni = pd.DataFrame()
        if s_tl:
            self.trenutni = pd.read_excel(xls, sheet_name=s_tl); self.trenutni.columns=[c.strip() for c in self.trenutni.columns]
            self.log(f"Trenutni lager: {len(self.trenutni)} redova")
        self.hist_df = pd.DataFrame()
        if s_hist:
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

    def _prepare_lookups(self):
        kp = self.prodaja[['ID KOMITENTA','id artikla','Naziv artikla','Grupa']].drop_duplicates()
        ks = self.startni[['ID KOMITENTA','id artikla','Naziv artikla','Grupa']].drop_duplicates()
        frames = [kp, ks]
        if self.has_history:
            kh = self.hist_df[['ID KOMITENTA','id artikla','Naziv artikla','Grupa']].drop_duplicates()
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
            # Count historical months per artikal
            for ida in self.hist_df['id artikla'].unique():
                sub=self.hist_df[self.hist_df['id artikla']==ida]
                self.hist_months_per_art[int(ida)]=sub[['Godina','Mesec']].drop_duplicates().shape[0]
            self.log(f"Istorijski prosek za {len(self.hist_dict)} kombinacija")
        # Count recent months per artikal
        self.recent_months_per_art={}
        for ida in self.prodaja['id artikla'].unique():
            sub=self.prodaja[self.prodaja['id artikla']==ida]
            self.recent_months_per_art[int(ida)]=sub[['Godina','Mesec']].drop_duplicates().shape[0]
        # Total months per artikal = hist + recent
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
            sales,oos,pocs=[],[],[]
            for god,mes in self.meseci_order:
                pv,lv,_=self.prodaja_dict.get((idk,ida,god,mes),(0,0,0)); lv=lv if not pd.isna(lv) else 0
                sales.append(pv); oos.append(1 if poc==0 else 0); pocs.append(poc); poc=lv
            ha=self.hist_dict.get((idk,ida),0)
            analysis.append({'idk':idk,'ida':ida,'sales':np.array(sales,dtype=float),'oos':np.array(oos),'poc':np.array(pocs,dtype=float),'ha':ha})
        preds={}
        for it in analysis:
            s,o,p=it['sales'],it['oos'],it['poc']; n=len(s); ha=it['ha']
            noos=s[o==0]
            if len(noos)>0 and noos.mean()>0:
                an=noos.mean(); adj=np.where(o==1,an,s)
                for m in range(n):
                    if o[m]==0 and p[m]>0 and p[m]<an*0.5: adj[m]=0.5*s[m]+0.5*an
            elif ha>0: adj=np.full(n,ha)
            else: adj=s.copy()
            if n>=2:
                lev=adj[0]; tr=(adj[-1]-adj[0])/max(n-1,1)
                for i in range(1,n):
                    nl=self.alpha*adj[i]+(1-self.alpha)*(lev+tr); nt=self.beta*(nl-lev)+(1-self.beta)*tr; lev,tr=nl,nt
                holt=lev+tr
            else: holt=adj[0]
            w=WMA_WEIGHTS[-n:] if n<=5 else WMA_WEIGHTS; w=w/w.sum()
            wma=np.dot(adj[-len(w):],w) if n>=3 else adj.mean()
            comb=0.5*holt+0.5*wma
            ma=adj.mean()
            if ma>0 and n>=3: comb*=(1+min((np.std(adj)/ma)*0.3,0.5))
            if ha>0 and comb>0: comb=(1-HIST_WEIGHT)*comb+HIST_WEIGHT*ha
            elif ha>0 and comb==0 and s.sum()==0: comb=ha*0.20
            # Full average: (hist_total + recent_total) / total_months
            ht=self.hist_total_dict.get((it['idk'],it['ida']),0)
            rt=float(s.sum())
            tm=self.total_months_per_art.get(it['ida'],n)
            full_avg=(ht+rt)/max(tm,1)
            preds[(it['idk'],it['ida'])]=(max(0,comb),full_avg)
        items=[{'k':k,'p':v[0],'a':v[1]} for k,v in preds.items()]; df_p=pd.DataFrame(items)
        df_p['pr']=df_p['p'].apply(lambda x: round(x) if x>=0.5 else (1 if x>0 else 0))
        for ida in df_p['k'].apply(lambda x:x[1]).unique():
            mask=df_p['k'].apply(lambda x:x[1]==ida); sub=df_p[mask]
            tgt=round(sub['p'].sum()); cur=sub['pr'].sum(); d=tgt-cur
            if d!=0:
                rem=sub['p']-np.floor(sub['p']); am=sub['p']>0
                if d>0:
                    for idx in rem[am].sort_values(ascending=False).index[:int(d)]: df_p.loc[idx,'pr']+=1
                elif d<0:
                    for idx in rem[am&(sub['pr']>0)].sort_values(ascending=True).index[:int(abs(d))]:
                        if df_p.loc[idx,'pr']>0: df_p.loc[idx,'pr']-=1
        df_p['ar']=df_p['a'].apply(lambda x:round(x))
        self.pred_dict={r['k']:(int(r['pr']),int(r['ar']),int(r['pr']-r['ar'])) for _,r in df_p.iterrows()}
        self.log(f"Predikcija: {sum(v[0] for v in self.pred_dict.values())} kom")

    def _merge_lager(self):
        for _,k in self.all_keys.iterrows():
            idk,ida=k['ID KOMITENTA'],k['id artikla']; pred,avg,razl=self.pred_dict.get((idk,ida),(0,0,0))
            lager=self.trenutni_dict.get((idk,ida),None)
            idx=self.df_monthly[(self.df_monthly['ID KOMITENTA']==idk)&(self.df_monthly['id artikla']==ida)].index
            if len(idx)>0:
                ix=idx[0]; self.df_monthly.loc[ix,'Predikcija']=pred; self.df_monthly.loc[ix,'Prosek']=avg; self.df_monthly.loc[ix,'Razlika']=razl
                if lager is not None: self.df_monthly.loc[ix,'Lager_danas']=lager
                else: self.df_monthly.loc[ix,'Lager_danas']=0
        for col in ['Predikcija','Prosek','Razlika','Lager_danas']:
            if col not in self.df_monthly.columns: self.df_monthly[col]=0
            self.df_monthly[col]=self.df_monthly[col].fillna(0).astype(int)

    def _compute_orders(self):
        self.df_result=self.df_monthly.copy()
        def p1(row):
            if row['ID KOMITENTA'] in self.excluded: return 0
            return max(int(row['Predikcija'])-int(row['Lager_danas']),0)
        def p2(row):
            if row['ID KOMITENTA'] in self.excluded: return 0
            return max(max(int(row['Predikcija'])-int(row['Lager_danas']),0),max(self.min_lager-int(row['Lager_danas']),0))
        self.df_result['Porudzbina_1']=self.df_result.apply(p1,axis=1).astype(int)
        self.df_result['Porudzbina_2']=self.df_result.apply(p2,axis=1).astype(int)

    def _apply_min_order(self):
        self.adjustments=[]
        for kid in sorted(self.df_result['ID KOMITENTA'].unique()):
            if kid in self.excluded: continue
            mask=self.df_result['ID KOMITENTA']==kid; total=self.df_result.loc[mask,'Porudzbina_2'].sum()
            if 1<=total<self.min_order:
                needed=self.min_order-total
                if total>=2:
                    cands=self.df_result.loc[mask&(self.df_result['Porudzbina_2']>0)].sort_values('Predikcija',ascending=False)
                    rem=int(needed)
                    for idx in cands.index:
                        if rem<=0: break
                        add=min(rem,2); self.df_result.loc[idx,'Porudzbina_2']+=add; rem-=add
                    if rem>0 and len(cands)>0: self.df_result.loc[cands.index[0],'Porudzbina_2']+=rem
                    self.adjustments.append((kid,f"{total}->{self.min_order}"))
                else:
                    self.df_result.loc[mask,'Porudzbina_2']=0; self.adjustments.append((kid,f"{total}->0"))


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
    cell=ws1.cell(1,ps,f'{engine.pred_label} - PREDIKCIJA'); cell.font=Font(bold=True,color='FFFFFF',name='Arial',size=10)
    cell.fill=pred_hdr; cell.alignment=ca
    for cc in range(ps,ps+3): ws1.cell(1,cc).border=tb; ws1.cell(1,cc).fill=pred_hdr
    for j,(sh,sfill) in enumerate(zip(['Predikcija','Prosek (svi mes.)','Razlika'],[sf_pred,sf_avg,sf_razl])):
        cell=ws1.cell(2,ps+j,sh); cell.font=sfnt; cell.fill=sfill; cell.border=tb; cell.alignment=caw
    os_c=ps+3
    ws1.merge_cells(start_row=1,end_row=1,start_column=os_c,end_column=os_c+2)
    cell=ws1.cell(1,os_c,f'PORUDZBINA - {engine.order_label}'); cell.font=Font(bold=True,color='FFFFFF',name='Arial',size=10)
    cell.fill=ord_hdr; cell.alignment=ca
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
            cell.alignment=Alignment(horizontal='center'); cell.border=tb
            if v>0: cell.fill=PatternFill('solid',fgColor='F3EEFA')
        for i,label in enumerate(ml):
            cb=month_start+i*SC
            for j,suf in enumerate(col_suf):
                cn=f'{label}{suf}'; v=row.get(cn,0)
                cell=ws1.cell(r,cb+j,int(v) if not pd.isna(v) else 0); cell.font=dfn; cell.alignment=Alignment(horizontal='center'); cell.border=tb
        for j,cn in enumerate(['Predikcija','Prosek','Razlika']):
            v=int(row.get(cn,0)); cell=ws1.cell(r,ps+j,v); cell.alignment=Alignment(horizontal='center'); cell.border=tb
            if cn=='Razlika':
                if v>0: cell.font=Font(name='Arial',size=9,color='006100',bold=True)
                elif v<0: cell.font=Font(name='Arial',size=9,color='9C0006',bold=True)
                else: cell.font=dfn
            else: cell.font=dfn
        for j,cn in enumerate(['Lager_danas','Porudzbina_1','Porudzbina_2']):
            v=int(row.get(cn,0)); cell=ws1.cell(r,os_c+j,v); cell.alignment=Alignment(horizontal='center'); cell.border=tb
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
    # Totali
    ws2=wb.create_sheet("Totali po mesecima")
    for c,h in enumerate(['Mesec','Promet (ulaz)','Prodaja','Stvarni povrat','Korekcija','Neto (Promet-Povrat)'],1):
        cell=ws2.cell(1,c,h); cell.font=Font(bold=True,color='FFFFFF',name='Arial',size=11); cell.fill=hf; cell.alignment=caw; cell.border=tb
    ro=2
    if engine.has_history:
        ws2.cell(ro,1,'Jan-Avg 2025 (UKUPNO)').font=Font(bold=True,name='Arial',size=10,color='6B3FA0')
        ws2.cell(ro,1).alignment=ca; ws2.cell(ro,1).border=tb
        cell=ws2.cell(ro,3,int(df['Total_JanAvg'].sum())); cell.font=Font(bold=True,name='Arial',size=10,color='6B3FA0')
        cell.fill=sf_hist; cell.alignment=ca; cell.border=tb; cell.number_format='#,##0'
        for c in [2,4,5,6]: ws2.cell(ro,c,'-').font=dfn; ws2.cell(ro,c).alignment=ca; ws2.cell(ro,c).border=tb
        ro+=2
    for ri,label in enumerate(ml,ro):
        ws2.cell(ri,1,label).font=Font(bold=True,name='Arial',size=10); ws2.cell(ri,1).alignment=ca; ws2.cell(ri,1).border=tb
        vals=[int(df[f'{label}_Promet'].sum()),int(df[f'{label}_Prodaja'].sum()),int(df[f'{label}_Povrat'].sum()),int(df[f'{label}_Korekcija'].sum())]
        vals.append(vals[0]-vals[2]); fills=[sf_prom,sf_prod,sf_pov,sf_kor,sf_poc]
        for c2,(v,f) in enumerate(zip(vals,fills),2):
            cell=ws2.cell(ri,c2,v); cell.font=Font(name='Arial',size=10); cell.fill=f; cell.alignment=ca; cell.border=tb; cell.number_format='#,##0'
    fr=ro+len(ml)+1
    ws2.cell(fr,1,f'PORUDZBINA {engine.order_label.upper()}').font=Font(bold=True,name='Arial',size=11,color='375623'); ws2.cell(fr,1).border=tb
    ir=[(f'Predikcija {engine.pred_label}',int(df['Predikcija'].sum()),sf_pred),('Prosek (svi meseci)',int(df['Prosek'].sum()),sf_avg),
        ('Trenutni lager',int(df['Lager_danas'].sum()),sf_lager),
        ('Porudzbina (osnovna)',int(df[~df['ID KOMITENTA'].isin(engine.excluded)]['Porudzbina_1'].sum()),sf_p1),
        (f'Porudzbina (min. {engine.min_lager})',int(df[~df['ID KOMITENTA'].isin(engine.excluded)]['Porudzbina_2'].sum()),sf_p2)]
    for i,(label,val,fill) in enumerate(ir,fr+1):
        ws2.cell(i,1,label).font=Font(bold=True,name='Arial',size=10); ws2.cell(i,1).alignment=ca; ws2.cell(i,1).border=tb
        cell=ws2.cell(i,2,val); cell.font=Font(bold=True,name='Arial',size=11); cell.fill=fill; cell.alignment=ca; cell.border=tb; cell.number_format='#,##0'
    ws2.column_dimensions['A'].width=32; ws2.column_dimensions['B'].width=18
    for c in 'CDEF': ws2.column_dimensions[c].width=18
    # O modelu
    ws3=wb.create_sheet("O modelu"); ws3.column_dimensions['A'].width=100
    info=["OPIS MODELA PREDIKCIJE I PORUDZBINE","",f"=== PREDIKCIJA ZA {engine.pred_label.upper()} ===","",
        "Model predvidja POTENCIJAL PRODAJE.","",f"  1. OOS korekcija",f"  2. Holt DES (alpha={engine.alpha}, beta={engine.beta}) + WMA",
        "  3. Parcijalni OOS blend","  4. Largest remainder zaokruzivanje"]
    if engine.has_history: info+=[f"  5. Istorijski podaci: {HIST_WEIGHT*100:.0f}% tezina","     - OOS objekti bez recentne prodaje koriste istoriju"]
    info+=["",f"=== PORUDZBINA ZA {engine.order_label.upper()} ===","",f"Osnovna: max(Pred-Lager, 0)",
        f"Min {engine.min_lager}: max(Pred-Lager, {engine.min_lager}-Lager, 0)",f"Min po objektu: >={engine.min_order} ili 0",
        f"Iskljuceni: {', '.join(str(x) for x in sorted(engine.excluded))}","",f"Generisano: {datetime.datetime.now().strftime('%d.%m.%Y. u %H:%M')}"]
    for i,line in enumerate(info,1):
        cell=ws3.cell(i,1,line)
        if i==1: cell.font=Font(bold=True,name='Arial',size=14,color='375623')
        elif '===' in line: cell.font=Font(bold=True,name='Arial',size=12,color='7030A0')
        else: cell.font=Font(name='Arial',size=10)
    buf=io.BytesIO(); wb.save(buf); buf.seek(0); return buf

# === STREAMLIT UI ===
import datetime

DEFAULT_EXCLUDED = "1023, 1027, 1034, 1043, 1057, 1060, 1061, 1076, 1315, 1347, 1349, 1359"

st.set_page_config(page_title="VAPE Porudzbine", page_icon="\U0001f4a8", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap');
    .stApp { background: linear-gradient(160deg, #fdf2f8 0%, #f5f0ff 40%, #eff6ff 100%); font-family: 'Poppins', sans-serif; }
    section[data-testid="stSidebar"] { background: linear-gradient(180deg, #7c3aed 0%, #a855f7 50%, #c084fc 100%) !important; }
    section[data-testid="stSidebar"] * { color: white !important; }
    section[data-testid="stSidebar"] input, section[data-testid="stSidebar"] textarea {
        background: rgba(255,255,255,0.15) !important; border: 1px solid rgba(255,255,255,0.3) !important;
        color: white !important; border-radius: 8px !important; }
    .metric-card { background: white; border-radius: 16px; padding: 20px 24px;
        box-shadow: 0 2px 12px rgba(124,58,237,0.08); border: 1px solid rgba(124,58,237,0.1); text-align: center; }
    .metric-value { font-size: 32px; font-weight: 700;
        background: linear-gradient(135deg, #7c3aed, #ec4899); -webkit-background-clip: text; -webkit-text-fill-color: transparent; }
    .metric-label { font-size: 13px; color: #888; margin-top: 4px; }
    .stButton > button { background: linear-gradient(135deg, #a855f7 0%, #ec4899 100%) !important;
        color: white !important; border: none !important; border-radius: 12px !important;
        padding: 12px 32px !important; font-weight: 600 !important; font-size: 16px !important;
        box-shadow: 0 4px 15px rgba(168,85,247,0.3) !important; }
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
</style>
""", unsafe_allow_html=True)

st.markdown("""<div class="header-banner"><div class="header-title">\U0001f4a8 VAPE PORUDZBINE</div>
    <div class="header-sub">Predikcija prodaje & Automatsko generisanje porudzbina</div></div>""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### Parametri")
    alpha = st.number_input("Alpha (nivo)", 0.0, 1.0, 0.4, 0.05)
    beta = st.number_input("Beta (trend)", 0.0, 1.0, 0.2, 0.05)
    min_lager = st.number_input("Min lager", 0, 20, 2)
    min_order = st.number_input("Min porudzbina po objektu", 0, 50, 5)
    st.markdown("---")
    st.markdown("### Iskljuceni komitenti")
    excluded_str = st.text_area("ID-evi razdvojeni zarezom", value=DEFAULT_EXCLUDED, height=100)

excluded = set()
for part in excluded_str.replace('\n', ',').split(','):
    p = part.strip()
    if p.isdigit(): excluded.add(int(p))

uploaded = st.file_uploader("Ucitaj Excel fajl (prodaja, startni lager, povrat, trenutni lager, prodaja pre septembra)", type=['xlsx','xls'])

if uploaded:
    file_bytes = uploaded.read()
    st.markdown(f'<div class="success-box">Fajl <strong>{uploaded.name}</strong> ucitan ({len(file_bytes)//1024} KB)</div>', unsafe_allow_html=True)
    st.markdown("")
    if st.button("POKRENI - Generisi porudzbinu", use_container_width=True):
        progress_bar = st.progress(0)
        try:
            engine = PredictionEngine(file_bytes, excluded, alpha, beta, min_lager, min_order)
            result = engine.run(progress_bar)
            st.markdown("---")
            tp = int(result['Predikcija'].sum()); tl = int(result['Lager_danas'].sum())
            t1 = int(result[~result['ID KOMITENTA'].isin(excluded)]['Porudzbina_1'].sum())
            t2 = int(result[~result['ID KOMITENTA'].isin(excluded)]['Porudzbina_2'].sum())
            no = result[result['Porudzbina_2']>0]['ID KOMITENTA'].nunique()
            c1,c2,c3,c4,c5 = st.columns(5)
            for col, val, lbl in [(c1,tp,'Predikcija'),(c2,tl,'Lager'),(c3,t1,'P1 osnovna'),(c4,t2,f'P2 min {min_lager}'),(c5,no,'Objekata')]:
                col.markdown(f'<div class="metric-card"><div class="metric-value">{val}</div><div class="metric-label">{lbl}</div></div>', unsafe_allow_html=True)
            st.markdown("")
            tab1, tab2, tab3 = st.tabs(["Porudzbina", "Sumarno", "Log"])
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
            with tab2:
                sm = result.groupby('ID KOMITENTA').agg(P1=('Porudzbina_1','sum'),P2=('Porudzbina_2','sum'),
                    Art=('Porudzbina_2', lambda x: int((x>0).sum()))).reset_index().sort_values('P2', ascending=False)
                st.dataframe(sm, use_container_width=True, height=400)
            with tab3:
                for msg in engine.logs: st.text(msg)
                if engine.adjustments:
                    st.markdown("**Korekcije:**")
                    for kid, note in engine.adjustments: st.text(f"  Komitent {kid}: {note}")
            st.markdown("---")
            excel_buf = create_excel(engine)
            fname = f"PORUDZBINA_{datetime.date.today().strftime('%Y%m%d')}.xlsx"
            st.download_button(f"Preuzmi Excel - {fname}", data=excel_buf, file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        except Exception as e:
            st.error(f"Greska: {str(e)}")
            import traceback; st.code(traceback.format_exc())
else:
    st.markdown("""<div style="text-align:center;padding:60px 20px;color:#aaa;">
        <div style="font-size:48px;margin-bottom:12px;">\U0001f4c2</div>
        <div style="font-size:16px;color:#888;">Ucitaj Excel fajl da pocnes</div>
        <div style="font-size:12px;color:#bbb;margin-top:8px;">Sheetovi: prodaja, startni lager, povrat, trenutni lager, prodaja pre septembra (opciono)</div>
    </div>""", unsafe_allow_html=True)
