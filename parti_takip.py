#!/usr/bin/env python3
# ============================================================
# PARTİ TAKİP SİSTEMİ - LOKAL VERSİYON
# Excel'den oku -> parti_takip.html + parti_veri.json üret
# ============================================================
import pandas as pd
import numpy as np
import re
import json
import sys
import os
import glob
from datetime import datetime

# ═══════════════════════════════════════════════
# KONFİGÜRASYON - KENDİ DOSYA YOLUNUZLA DEĞİŞTİRİN
# ═══════════════════════════════════════════════
EXCEL_PATH = "veri.xlsx"
OUTPUT_HTML = "index.html"
OUTPUT_JSON = "parti_veri.json"

BEKLEME_UYARI_GUN = 7
BEKLEME_KRITIK_GUN = 14
FILTRE_SON_ASAMA = None
OZEL_ASAMA_SIRASI = []

TR_MAP = str.maketrans({"ç":"c","Ç":"c","ğ":"g","Ğ":"g","ı":"i","I":"i","İ":"i","ö":"o","Ö":"o","ş":"s","Ş":"s","ü":"u","Ü":"u"})

def norm(x):
    if pd.isna(x): return ""
    return re.sub(r"\s+"," ",str(x).strip().translate(TR_MAP).lower())

def cs(x):
    if pd.isna(x): return ""
    if isinstance(x,float) and x.is_integer(): return str(int(x))
    s=str(x).strip(); return "" if s.lower()=="nan" else s

def uj(series):
    v=[]
    for x in series:
        s=cs(x)
        if s and s not in v: v.append(s)
    return " | ".join(v)

def tn(x):
    if pd.isna(x): return 0.0
    if isinstance(x,(int,float,np.integer,np.floating)): return float(x)
    s=str(x).replace(",","."); m=re.search(r"-?\d+(?:\.\d+)?",s)
    return float(m.group()) if m else 0.0

def parse_bekleme_gun(x):
    if pd.isna(x): return np.nan
    if isinstance(x,pd.Timedelta): return round(x.total_seconds()/86400,4)
    if isinstance(x,(int,float,np.integer,np.floating)): return round(float(x),4)
    s=str(x).strip()
    if not s: return np.nan
    s2=re.sub(r"(GÜN|gün|Gun|gun)","Gün",s,flags=re.I)
    m=re.search(r"(\d+)\s+Gün\s*,?\s*(\d{1,2}):(\d{1,2}):(\d{1,2})",s2,re.I)
    if m: return int(m.group(1))+int(m.group(2))/24+int(m.group(3))/1440+int(m.group(4))/86400
    m=re.search(r"(\d+)\s+Gün",s2,re.I)
    if m: return float(m.group(1))
    m=re.search(r"^(\d{1,2}):(\d{1,2}):(\d{1,2})",s)
    if m: return int(m.group(1))/24+int(m.group(2))/1440+int(m.group(3))/86400
    return np.nan

def format_bekleme_detay(x):
    if pd.isna(x): return ""
    if isinstance(x,str):
        s=x.strip()
        m=re.search(r"(\d+)\s*(?:Gün|Gun|gün|gun|GÜN)\s*,?\s*(\d{1,2}):(\d{2}):(\d{2})",s,re.I)
        if m: return f"{int(m.group(1))} Gün, {int(m.group(2)):02d}:{m.group(3)}:{m.group(4)}"
        m=re.search(r"(\d+)\s*(?:Gün|Gun|gün|gun|GÜN)",s,re.I)
        if m: return f"{int(m.group(1))} Gün, 00:00:00"
        m=re.search(r"^(\d{1,2}):(\d{2}):(\d{2})",s)
        if m: return f"0 Gün, {int(m.group(1)):02d}:{m.group(2)}:{m.group(3)}"
        return s
    if isinstance(x,pd.Timedelta):
        ts2=int(x.total_seconds())
        if ts2<0: ts2=0
        g=ts2//86400;k=ts2%86400;sa=k//3600;dk=(k%3600)//60;sn=k%60
        return f"{g} Gün, {sa:02d}:{dk:02d}:{sn:02d}"
    if isinstance(x,(int,float,np.integer,np.floating)):
        v=float(x)
        if np.isnan(v): return ""
        g=int(v);ks=int((v-g)*86400)
        if ks<0: ks=0
        sa=ks//3600;dk=(ks%3600)//60;sn=ks%60
        return f"{g} Gün, {sa:02d}:{dk:02d}:{sn:02d}"
    return str(x)

def fc(df,cands):
    nc={c:norm(c) for c in df.columns}
    for c in cands:
        n=norm(c)
        for r,nr in nc.items():
            if nr==n: return r
    for c in cands:
        n=norm(c)
        for r,nr in nc.items():
            if n in nr or nr in n: return r
    return None

def kisalt_firma(firma):
    if not firma: return ""
    f=firma.strip().upper()
    for p in [r"\sVE\s+TEKS\w*\s+SAN\w*\s+VE\s+TİC\w*\.?\sA\.?Ş\.?",r"\sVE\s+TİC\w*\.?\sA\.?Ş\.?",r"\sSAN\w*\.?\sVE\s+TİC\w*\.?\sA\.?Ş\.?",r"\sSAN\w*\.?\sTİC\w*\.?\sA\.?Ş\.?",r"\sTEKST?\w*\.?\sSAN\w*\.?",r"\sSAN\w*\.?\sA\.?Ş\.?",r"\sTİC\w*\.?\sA\.?Ş\.?",r"\sA\.?Ş\.?\s*",r"\s*LTD\.?\s*ŞTİ\.?\s*",r"\sLTD\.?\s*",r"\s*ŞTİ\.?\s*"]:
        f=re.sub(p,"",f,flags=re.I)
    f=re.sub(r"\s+"," ",f).strip().rstrip("., ")
    if len(f)>18: f=f.split()[0] if f.split() else f
    return f

def kisalt_siparis_no(sip):
    if not sip: return ""
    s=sip.strip();s=re.sub(r"[/-]?\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}","",s)
    parts=[p.strip() for p in re.split(r"[/]",s) if p.strip()]
    return parts[0] if parts else ""

def kisalt_musteri_sip(msip):
    if not msip: return ""
    s=msip.strip();s=re.sub(r"\d{4}\d{4}\s*","",s);return re.sub(r"\s+"," ",s).strip()

def kisalt_ham_adi(ham):
    if not ham: return ""
    ham=re.sub(r"\(.*?\)","",ham).strip();ham=re.sub(r"(\w+\d+)\s+(\d+)",r"\1/\2",ham);return ham

def kisalt_recete(recete):
    if not recete: return ""
    recete=re.sub(r"\(.*?\)","",recete).strip()
    k={"BOYAMA":"BY","BOYA":"BY","AĞARTMA":"AĞR","AGARTMA":"AĞR","YIKAMA":"YIK","KASARLAMA":"KSR","APRE":"APR","BASKILI":"BSK","BASKI":"BSK","SİLİKON":"SLK","SILIKON":"SLK","YUMUŞATMA":"YMŞ","YUMUSATMA":"YMŞ","MERSER":"MRS","SANFOR":"SNF","KALENDİR":"KLD","KALENDIR":"KLD"}
    kl=recete.strip().upper().split();s=[]
    for w in kl:
        if re.search(r"\d",w) or "/" in w or len(w)<=2: s.append(w)
        elif w in k: s.append(k[w])
        elif len(w)>5: s.append(w[:4])
        else: s.append(w)
    return " ".join(s)

def bul_bir_sonraki_asama(uretim_asamalari, sonra_yapilacak, son_yapilan=""):
    if pd.isna(uretim_asamalari) or pd.isna(sonra_yapilacak): return ""
    ua_str=str(uretim_asamalari).strip();sy_str=str(sonra_yapilacak).strip();last_str=str(son_yapilan).strip()
    if not ua_str or not sy_str: return ""
    asamalar=[a.strip() for a in ua_str.split(",") if a.strip()]
    if not asamalar: return ""
    def eslesiyor_mu(s1,s2):
        if not s1 or not s2: return False
        c1=re.sub(r"[^\w\s]"," ",norm(s1)).strip();c2=re.sub(r"[^\w\s]"," ",norm(s2)).strip()
        if c1==c2: return True
        if c1.replace(" ","") in c2.replace(" ","") or c2.replace(" ","") in c1.replace(" ",""): return True
        w1=c1.split();w2=c2.split()
        if len(w1)==len(w2) and len(w1)>0: return all(k1.startswith(k2) or k2.startswith(k1) for k1,k2 in zip(w1,w2))
        return False
    hedef_indeksler=[i for i,a in enumerate(asamalar) if eslesiyor_mu(a,sy_str)]
    son_indeksler=[i for i,a in enumerate(asamalar) if eslesiyor_mu(a,last_str)] if last_str else []
    bulunan_idx=-1
    if son_indeksler and hedef_indeksler:
        en_son=son_indeksler[-1]
        for h_idx in hedef_indeksler:
            if h_idx>=en_son: bulunan_idx=h_idx;break
        if bulunan_idx==-1: bulunan_idx=hedef_indeksler[-1]
    elif hedef_indeksler: bulunan_idx=hedef_indeksler[0]
    if bulunan_idx==-1: return ""
    sonraki_idx=bulunan_idx+1
    return asamalar[sonraki_idx] if sonraki_idx<len(asamalar) else ""

def build_line1(r):
    fk=kisalt_firma(r["Firma Adı"]);sk=kisalt_siparis_no(r["Sipariş No"]);mk=kisalt_musteri_sip(r["Müşteri Sipariş"])
    parts=[]
    if fk: parts.append(fk)
    if sk and sk.upper()!=fk.upper(): parts.append(sk)
    if mk: parts.append(mk)
    return " / ".join(parts)

def build_line2(r):
    h=kisalt_ham_adi(r["Ham Adı"]) if r["Ham Adı"] else "";rc=kisalt_recete(r["Reçete Adı"]) if r["Reçete Adı"] else ""
    return " — ".join([x for x in [h,rc] if x])

def process_excel(input_path):
    print(f"Excel okunuyor: {input_path}")
    xls=pd.ExcelFile(input_path);frames=[]
    for sh in xls.sheet_names:
        t=pd.read_excel(input_path,sheet_name=sh)
        if t.empty: continue
        if fc(t,["Parti No"]) and fc(t,["SONRAKİ","Sonra Yapılacak Aşama","Sonraki Aşama"]): frames.append(t)
    if not frames: raise ValueError("Uygun sayfa bulunamadı.")
    df_raw=pd.concat(frames,ignore_index=True)

    col_parti=fc(df_raw,["Parti No"])
    col_next1=fc(df_raw,["SONRAKİ","Sonra Yapılacak Aşama","Sonraki Aşama"])
    col_musteri=fc(df_raw,["MUSTERISIPARISNO","Müşteri Sipariş No","Müşteri Sipariş"])
    col_siparis=fc(df_raw,["Sipariş No"])
    col_firma=fc(df_raw,["Firma Adı"])
    col_kilo=fc(df_raw,["Kilo","Kalan Brüt Kilo"])
    col_wait=fc(df_raw,["Çıkıştan Sonra Geçen Süre","Bekleme Notu"])
    col_son=fc(df_raw,["Son Aşama","Son Yapılan Aşama"])
    col_cikis=fc(df_raw,["Çıkış Tarihi","Son Hareket Tarihi"])
    col_ham=fc(df_raw,["Ham Adı","Ham Ad","Hammadde","Ham Madde Adı","HamAdi"])
    col_recete=fc(df_raw,["Renk Adı","Reçete Adı","Reçete Kodu","Reçete","Recete","ReceteAdi"])
    col_uretim=fc(df_raw,["Üretim Aşamaları","Uretim Asamalari","Üretim Aşaması","Aşamalar","Asamalar","Üretim Sırası","Uretim Sirasi","İş Akışı","Is Akisi","Proses Sırası"])

    df=df_raw.copy()
    df["Parti No"]=df[col_parti].apply(cs)
    df["Hedef Aşama"]=df[col_next1].apply(cs) if col_next1 else ""
    df["Son Yapılan"]=df[col_son].apply(cs) if col_son else ""
    df["Müşteri Sipariş"]=df[col_musteri].apply(cs) if col_musteri else ""
    df["Sipariş No"]=df[col_siparis].apply(cs) if col_siparis else ""
    df["Firma Adı"]=df[col_firma].apply(cs) if col_firma else ""
    df["Kilo"]=df[col_kilo].apply(tn) if col_kilo else 0.0
    df["Ham Adı"]=df[col_ham].apply(cs) if col_ham else ""
    df["Reçete Adı"]=df[col_recete].apply(cs) if col_recete else ""

    if col_uretim and col_next1:
        df["Üretim Aşamaları"]=df[col_uretim].apply(cs)
        df["Bir Sonraki Aşama"]=df.apply(lambda r: bul_bir_sonraki_asama(r["Üretim Aşamaları"],r["Hedef Aşama"],r["Son Yapılan"]),axis=1)
    else:
        df["Bir Sonraki Aşama"]="";df["Üretim Aşamaları"]=""

    if col_wait:
        df["Bekleme Gün"]=df[col_wait].apply(parse_bekleme_gun)
        df["Bekleme Detay"]=df[col_wait].apply(format_bekleme_detay)
    elif col_cikis:
        cikis_dt=pd.to_datetime(df[col_cikis],errors="coerce",dayfirst=True)
        delta=pd.Timestamp.today().normalize()-cikis_dt.dt.normalize()
        df["Bekleme Gün"]=delta.dt.days.astype(float)
        df["Bekleme Detay"]=delta.apply(format_bekleme_detay)
    else:
        df["Bekleme Gün"]=np.nan;df["Bekleme Detay"]=""

    if FILTRE_SON_ASAMA and col_son: df=df[df[col_son].apply(lambda x: norm(x)==norm(FILTRE_SON_ASAMA))].copy()
    df=df[(df["Parti No"]!="")&(df["Hedef Aşama"]!="")].copy()

    df["_line1"]=df.apply(build_line1,axis=1);df["_line2"]=df.apply(build_line2,axis=1)

    def agg_bekleme_detay(g):
        v=g.dropna(subset=["Bekleme Gün"])
        if v.empty: return ""
        return v.loc[v["Bekleme Gün"].idxmax(),"Bekleme Detay"]

    ozet_parts=[]
    for (ha,pn),grp in df.groupby(["Hedef Aşama","Parti No"]):
        ozet_parts.append({"Hedef Aşama":ha,"Parti No":pn,"_line1":uj(grp["_line1"]),"_line2":uj(grp["_line2"]),
        "Firma Adı":uj(grp["Firma Adı"]),"Müşteri Sipariş":uj(grp["Müşteri Sipariş"]),
        "Sipariş No":uj(grp["Sipariş No"]),"Ham Adı":uj(grp["Ham Adı"]),"Reçete Adı":uj(grp["Reçete Adı"]),
        "Kilo":grp["Kilo"].sum(),"Bekleme Gün":grp["Bekleme Gün"].max(),"Bekleme Detay":agg_bekleme_detay(grp),
        "Bir Sonraki Aşama":uj(grp["Bir Sonraki Aşama"]),"Üretim Aşamaları":uj(grp["Üretim Aşamaları"])})

    ozet=pd.DataFrame(ozet_parts)
    vs=[]
    for x in df["Hedef Aşama"]:
        if x not in vs: vs.append(x)
    asama_sirasi=[]
    for x in OZEL_ASAMA_SIRASI+vs:
        if x and x not in asama_sirasi: asama_sirasi.append(x)
    bir_sonraki_sirasi=[]
    for x in df["Bir Sonraki Aşama"]:
        if x and x not in bir_sonraki_sirasi: bir_sonraki_sirasi.append(x)

    rm={x:i for i,x in enumerate(asama_sirasi)}
    ozet["__r"]=ozet["Hedef Aşama"].map(rm).fillna(9999)
    ozet=ozet.sort_values(["__r","Hedef Aşama","Bekleme Gün","Kilo","Parti No"],ascending=[True,True,False,False,True]).drop(columns="__r")

    ts=datetime.now().strftime("%d.%m.%Y %H:%M");gk=ozet["Kilo"].sum()

    asama_data={};asama_order=[]
    for a in asama_sirasi:
        alt=ozet[ozet["Hedef Aşama"]==a].copy()
        alt=alt.sort_values(["Bekleme Gün","Kilo","Parti No"],ascending=[False,False,True],na_position="last")
        kartlar=[]
        for _,r in alt.iterrows():
            bg=r["Bekleme Gün"]
            seviye="kritik" if pd.notna(bg) and bg>=BEKLEME_KRITIK_GUN else "uyari" if pd.notna(bg) and bg>=BEKLEME_UYARI_GUN else "normal"
            kartlar.append({"parti":r["Parti No"],"asama":r["Hedef Aşama"],"bir_sonraki":r.get("Bir Sonraki Aşama",""),
            "uretim_asamalari":r.get("Üretim Aşamaları",""),"firma":r.get("Firma Adı",""),
            "musteri_siparis":r.get("Müşteri Sipariş",""),"siparis_no":r.get("Sipariş No",""),
            "ham_adi":r.get("Ham Adı",""),"recete_adi":r.get("Reçete Adı",""),
            "line1":r["_line1"],"line2":r["_line2"],"kilo":float(r["Kilo"]),
            "bekleme":r.get("Bekleme Detay",""),"bekleme_gun":float(bg) if pd.notna(bg) else 0,"seviye":seviye})
        if kartlar:
            asama_data[a]={"kartlar":kartlar,"toplam_kg":float(alt["Kilo"].sum()),"parti_sayisi":len(kartlar)}
            asama_order.append(a)

    return asama_data, asama_order, bir_sonraki_sirasi, ts, gk

def build_json_file(asama_data, asama_order, bir_sonraki_sirasi):
    return json.dumps({
        "DATA": asama_data,
        "ASAMA_ORDER": asama_order,
        "BSA_ORDER": bir_sonraki_sirasi,
        "timestamp": datetime.now().isoformat()
    }, ensure_ascii=False)

def main():
    if len(sys.argv) > 1:
        excel_path = sys.argv[1]
    else:
        excel_path = EXCEL_PATH

    if not os.path.exists(excel_path):
        # Alternatif uzantıları dene
        xlsx_files = glob.glob("veri.Xlsx") + glob.glob("veri.xlsx") + glob.glob("VERI.XLSX") + glob.glob("veri.*")
        if xlsx_files:
            excel_path = xlsx_files[0]
            print(f"Excel bulundu: {excel_path}")
        else:
            print(f"HATA: Excel dosyası bulunamadı: {excel_path}")
            print("Mevcut .xlsx dosyaları:")
            for f in glob.glob("*"):
                print(f"  {f}")
            sys.exit(1)

    asama_data, asama_order, bir_sonraki_sirasi, ts, gk = process_excel(excel_path)
    len_ozet = sum(v["parti_sayisi"] for v in asama_data.values())

    data_json = json.dumps(asama_data, ensure_ascii=False).replace("</", "<\\/")
    order_json = json.dumps(asama_order, ensure_ascii=False).replace("</", "<\\/")
    bsa_order_json = json.dumps(bir_sonraki_sirasi, ensure_ascii=False).replace("</", "<\\/")

    with open("template.html", "r", encoding="utf-8") as f:
        html_template = f.read()

    html = html_template.replace("__DATA_JSON__", data_json)
    html = html.replace("__ORDER_JSON__", order_json)
    html = html.replace("__BSA_ORDER_JSON__", bsa_order_json)
    html = html.replace("__TS__", ts)
    html = html.replace("__LEN_OZET__", str(len_ozet))
    html = html.replace("__GK__", f"{gk:,.0f}")

    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)

    json_content = build_json_file(asama_data, asama_order, bir_sonraki_sirasi)
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        f.write(json_content)

    print(f"✅ {OUTPUT_HTML} + {OUTPUT_JSON} oluşturuldu")
    print(f"📌 {len_ozet} parti • {len(asama_order)} aşama • {gk:,.0f} kg")

if __name__ == "__main__":
    main()