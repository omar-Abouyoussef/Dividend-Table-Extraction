
import streamlit as st
import pdfplumber, pandas as pd, numpy as np, re, math, bisect
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import datetime
from difflib import SequenceMatcher
from io import BytesIO
from scipy.optimize import linear_sum_assignment
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="EGX Report Extractor", page_icon="📊", layout="wide")
st.title("EGX Report Extractor")
st.caption("Extract EGX daily market data and export matched tickers to Excel.")

# ----------------------------- SETTINGS ---------------------------------------
TABLE_TITLE = "Main Market Stocks Indicators"
NAME_HEADER = "Stock Name"
WANTED_COLUMNS = ["No. Listed Shares", "Market Capital in EGP", "Paid Capital in EGP",
                  "DY%", "EGP Net Profit", "Net Profit Date", "Last Coupon", "Pay Date"]

MANUAL_ALIASES = {
    "T M G Holding": "TMGH",
    "Egyptian Gulf Marseilia For Real Estate Investment": "MAAL",
    "Natural Gas & Mining Project (Egypt Gas)": "EGAS",
    "MNT Tech Holding For Financial Investments": "HALN",
    "Engineering Industries (ICON)": "PACH",
}

# --------------------------- PDF EXTRACTION -----------------------------------
def norm(s):
    return re.sub(r"[^a-z0-9%#]+", "", str(s).lower())

def lines(words):
    out = []
    for w in sorted(words, key=lambda w: (w["top"], w["x0"])):
        if out and abs(w["top"] - out[-1][0]) <= 0.5 * (w["bottom"] - w["top"]):
            out[-1][1].append(w)
        else:
            out.append([w["top"], [w]])
    return [(t, sorted(ws, key=lambda w: w["x0"])) for t, ws in out]

def text(words):
    return " ".join(" ".join(w["text"] for w in ws) for _, ws in lines(words))

def cx(w):
    return (w["x0"] + w["x1"]) / 2

@dataclass
class Band:
    top: float
    bottom: float
    edges: list
    labels: list
    @property
    def signature(self):
        return tuple(norm(l) for l in self.labels)
    def col_of(self, x):
        if not self.edges[0] <= x < self.edges[-1]:
            return None
        return bisect.bisect_right(self.edges, x) - 1

def find_bands(page, words, name_header):
    first = name_header.split()[0].lower()
    seeds = [w for w in words if w["text"].lower() == first]
    groups = defaultdict(dict)
    for r in page.rects:
        groups[(round(r["top"], 1), round(r["bottom"], 1))][round(r["x0"], 1)] = r
    bands = []
    for (top, bottom), cells in groups.items():
        if len(cells) < 3:
            continue
        rs = [cells[k] for k in sorted(cells)]
        if not any(r["x0"] <= cx(w) < r["x1"] and top-1 <= (w["top"]+w["bottom"])/2 <= bottom+1 for r in rs for w in seeds):
            continue
        labels = [text([w for w in words if r["x0"] <= cx(w) < r["x1"] and r["top"]-1 <= (w["top"]+w["bottom"])/2 <= r["bottom"]+1]) for r in rs]
        if norm(name_header) in [norm(l) for l in labels]:
            bands.append(Band(top, bottom, [r["x0"] for r in rs]+[rs[-1]["x1"]], labels))
    return sorted(bands, key=lambda b: b.top)

def row_bounds(page, band, end_y):
    rs = [r for r in page.rects if abs(r["x0"]-band.edges[0]) < 1 and abs(r["x1"]-band.edges[1]) < 1 and r["top"] >= band.bottom-1 and (end_y is None or r["top"] < end_y)]
    if len(rs) < 2:
        return None
    b = sorted({round(band.bottom,1)} | {round(v,1) for r in rs for v in (r["top"],r["bottom"])})
    b = [v for i,v in enumerate(b) if i == 0 or v-b[i-1] > 1]
    b.append(b[-1] + float(np.median(np.diff(b))))
    return b

def heading_above(words, band):
    return any(ws[0]["text"].startswith("-") and t < band.top for t,ws in lines(words))

def page_rows(page, words, band, end_y):
    body = [w for w in words if w["top"] >= band.bottom-1 and (end_y is None or w["top"] < end_y)]
    anchors = sorted([w for w in body if band.col_of(cx(w)) == 0 and re.fullmatch(r"\d{1,5}",w["text"])], key=lambda w:w["top"])
    if not anchors:
        return []
    ys = [a["top"] for a in anchors]
    pitch = float(np.median(np.diff(ys))) if len(ys)>1 else band.bottom-band.top
    bounds = row_bounds(page, band, end_y)
    interval = None
    if bounds:
        def interval(w):
            k = bisect.bisect_right(bounds,(w["top"]+w["bottom"])/2)-1
            return k if 0<=k<len(bounds)-1 else None
        per = [interval(a) for a in anchors]
        if None not in per and len(set(per)) == len(per):
            slot = {k:i for i,k in enumerate(per)}
        else:
            interval = None
    def row_of(w):
        if interval is not None:
            return slot.get(interval(w))
        i = int(np.argmin([abs(w["top"]-y) for y in ys]))
        dy = w["top"]-ys[i]
        if (i==0 and dy<-.5*pitch) or (i==len(ys)-1 and dy>.5*pitch):
            return None
        return i
    cells = [defaultdict(list) for _ in anchors]
    for w in body:
        c, i = band.col_of(cx(w)), row_of(w)
        if c is not None and c != 0 and i is not None:
            cells[i][c].append(w)
    rows = []
    for a,cell in zip(anchors,cells):
        row = {band.labels[0] or "#":int(a["text"])}
        for c in range(1,len(band.labels)):
            row[band.labels[c]] = text(cell.get(c,[])) or None
        rows.append(row)
    return rows

def extract_table(pdf, title=TABLE_TITLE, name_header=NAME_HEADER):
    rows, signature, collecting = [], None, False
    with pdfplumber.open(pdf) as doc:
        for pno,page in enumerate(doc.pages,1):
            words = page.extract_words()
            bands = find_bands(page,words,name_header)
            if not collecting:
                title_top = next((t for t,ws in lines(words) if title.lower() in text(ws).lower()),None)
                cand = [b for b in bands if title_top is not None and b.top>=title_top-1]
                if not cand:
                    page.flush_cache(); continue
                band,signature,collecting = cand[0],cand[0].signature,True
            else:
                if not bands or bands[0].signature!=signature or heading_above(words,bands[0]):
                    break
                band = bands[0]
            later = [b for b in bands if b.top>band.top]
            end_y = later[0].top if later else None
            for r in page_rows(page,words,band,end_y):
                r["pdf_page"] = pno
                rows.append(r)
            page.flush_cache()
            if later:
                break
    if not rows:
        raise ValueError(f'Table "{title}" with "{name_header}" was not found.')
    df = pd.DataFrame(rows)
    key = df.columns[0]
    if df[key].tolist()!=list(range(1,len(df)+1)):
        raise ValueError("Row numbers are not sequential. Extraction missed or merged rows.")
    return df

# ----------------------------- MATCHING ---------------------------------------
STOP = {"co","company","companies","the","for","of","and","sae","in","al","el","egypt","egyptian","group","corp","corporation","ltd","plc","inc","sa","s","a","e"}
R = lambda a,b: SequenceMatcher(None,a,b).ratio()

def toks(s, drop=True):
    s = re.sub(r"[^a-z0-9\s]"," ",str(s).lower().replace("&"," and "))
    out,buf=[], ""
    for w in s.split():
        if len(w)==1 and w.isalpha():
            buf+=w
        else:
            if buf: out.append(buf); buf=""
            out.append(w)
    if buf: out.append(buf)
    return ([w for w in out if w not in STOP] or out) if drop else out

class Matcher:
    def __init__(self,a_names,b_names):
        corpus=[toks(x) for x in list(a_names)+list(b_names)]
        self.n=len(corpus); self.df=Counter(w for t in corpus for w in set(t))
    def idf(self,w):
        return math.log((self.n+1)/(self.df.get(w,0)+1))+.3
    def tok(self,w,tb):
        best=0
        for v in tb:
            r=R(w,v)
            if len(w)>=3 and len(v)>=3 and (v.startswith(w) or w.startswith(v)):
                r=max(r,.9)
            best=max(best,r)
        return best
    def soft(self,ta,tb):
        def d(x,y):
            tot=sum(self.idf(w) for w in x)
            return sum(self.idf(w)*self.tok(w,y) for w in x)/tot if tot else 0
        a,b=d(ta,tb),d(tb,ta)
        return 100*(2*a*b/(a+b) if a+b else 0)
    def score(self,a,b):
        ta,tb=toks(a),toks(b)
        ss=self.soft(ta,tb)
        sc=100*max(R(" ".join(sorted(ta))," ".join(sorted(tb))),R("".join(ta),"".join(tb)),R(" ".join(toks(a,False))," ".join(toks(b,False))))
        return max(ss,sc) if sc>=95 else .6*ss+.4*sc

def load_companies(file):
    c=pd.read_csv(file).fillna("")
    t=next((k for k in c.columns if "ticker" in k.lower() or "symbol" in k.lower()),c.columns[0])
    n=next((k for k in c.columns if "name" in k.lower() and k!=t),c.columns[1])
    c=c[[t,n]].rename(columns={t:"ticker",n:"company name"})
    c["ticker"]=c["ticker"].astype(str).str.strip().str.upper()
    c["company name"]=c["company name"].astype(str).str.strip()
    return c[(c["ticker"]!="")&(c["company name"]!="")].drop_duplicates().reset_index(drop=True)

def match_tickers(df,companies,name_col=NAME_HEADER,aliases=None):
    aliases={str(k).strip().lower():(v or None) for k,v in (aliases or {}).items()}
    names,tick=companies["company name"].tolist(),companies["ticker"].tolist()
    t2i={t:i for i,t in enumerate(tick)}
    m=Matcher(df[name_col],names)
    S=np.array([[m.score(a,b) for b in names] for a in df[name_col]])
    ri,ci=linear_sum_assignment(-S)
    assign=dict(zip(ri,ci)); res=[]
    for i,name in enumerate(df[name_col]):
        j=assign[i]; top=np.argsort(-S[i])[:2]
        ticker,matched,score,status=tick[j],names[j],float(S[i,j]),None
        key=str(name).strip().lower()
        if key in aliases:
            if aliases[key] is None or aliases[key] not in t2i:
                ticker,matched,status=None,None,"No match"
            else:
                ticker=aliases[key]; j=t2i[ticker]; matched,score,status=names[j],float(S[i,j]),"Manual alias"
        if status is None:
            status="High" if score>=90 else "Medium" if score>=75 else "Low"
        res.append({"Ticker":ticker,"Matched Company (CSV)":matched,"Match Score":round(score,1),"Match Status":status,
                    "Closest (unconstrained)":f"{tick[top[0]]} ({S[i,top[0]]:.0f})","Runner-up":f"{tick[top[1]]} ({S[i,top[1]]:.0f})"})
    return pd.concat([df.reset_index(drop=True),pd.DataFrame(res)],axis=1)

# -------------------------- EXCEL EXPORT --------------------------------------
NUM=re.compile(r"^-?\d[\d,]*(\.\d+)?$")
DATE=re.compile(r"^\d{2}/\d{2}/\d{4}$")
PCT=re.compile(r"^-?[\d,]*\.?\d+%$")
AMT=re.compile(r"^(-?[\d,]*\.?\d+)\s+([A-Za-z]+)$")
F=lambda s:float(str(s).replace(",","").replace("%",""))

def clean_column(name,s):
    v=s.dropna().astype(str)
    if v.empty: return [(name,s,"@")]
    if v.map(lambda x:bool(NUM.match(x))).all():
        dec=max((len(x.split(".")[1]) if "." in x else 0) for x in v)
        return [(name,s.map(lambda x:None if pd.isna(x) else F(x)),"#,##0"+("."+"0"*dec if dec else ""))]
    if v.map(lambda x:bool(DATE.match(x))).all():
        return [(name,s.map(lambda x:None if pd.isna(x) else datetime.strptime(x,"%d/%m/%Y")),"dd/mm/yyyy")]
    if v.map(lambda x:bool(PCT.match(x))).all():
        return [(name,s.map(lambda x:None if pd.isna(x) else F(x)/100),"0.0%")]
    if v.map(lambda x:bool(AMT.match(x))).all():
        parts=s.map(lambda x:None if pd.isna(x) else AMT.match(x).groups())
        dec=max(len(AMT.match(x).group(1).split(".")[1]) if "." in AMT.match(x).group(1) else 0 for x in v)
        return [(name,parts.map(lambda p:None if p is None else F(p[0])),"0."+"0"*dec if dec else "0"),
                (name+" Currency",parts.map(lambda p:None if p is None else p[1]),"@")]
    return [(name,s,"@")]

def select_export(matched,wanted):
    by_norm={norm(c):c for c in matched.columns}
    rank=matched.columns[0]
    order=[rank,"Ticker",NAME_HEADER]
    data={c:matched[c] for c in order}
    fmts={rank:"0","Match Score":"0.0"}
    missing=[]
    for w in wanted:
        src=by_norm.get(norm(w))
        if src is None:
            missing.append(w); continue
        for nm,ser,fmt in clean_column(src,matched[src]):
            data[nm]=ser; fmts[nm]=fmt
    # Keep only requested export columns in exact order
    desired=[rank,"Ticker",NAME_HEADER]+wanted
    if "Last Coupon Currency" in data:
        desired.insert(desired.index("Pay Date"),"Last Coupon Currency")
    out=pd.DataFrame(data)[desired]
    out=out.rename(columns={rank:"#"})
    return out,fmts,missing

FILLS={"High":"E2EFDA","Medium":"FFF2CC","Low":"FCE4D6","Manual alias":"DDEBF7","No match":"F8CBAD"}

def write_sheet(ws,frame,fmts):
    hf=PatternFill("solid",start_color="1F4E78")
    hfont=Font(bold=True,color="FFFFFF",size=10)
    line=Border(bottom=Side(style="thin",color="D9D9D9"))
    for j,col in enumerate(frame.columns,1):
        c=ws.cell(1,j,col); c.font=hfont; c.fill=hf
        c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
        longest=max([len(str(col))]+[len(str(x)) for x in frame[col].head(300) if x is not None])
        ws.column_dimensions[get_column_letter(j)].width=min(max(10,longest+2),42)
    ws.row_dimensions[1].height=32
    for i,rec in enumerate(frame.to_dict("records"),2):
        for j,col in enumerate(frame.columns,1):
            v=rec[col]
            if v is None or v is pd.NaT or (isinstance(v,float) and np.isnan(v)): v=None
            elif isinstance(v,pd.Timestamp): v=v.to_pydatetime()
            c=ws.cell(i,j,v); c.font=Font(name="Arial",size=10); c.border=line
            if col in fmts: c.number_format=fmts[col]
            if j<=3: c.alignment=Alignment(horizontal="center")
    ws.freeze_panes="D2"; ws.auto_filter.ref=f"A1:{get_column_letter(len(frame.columns))}{len(frame)+1}"

def excel_bytes(data,fmts,matched):
    wb=Workbook(); ws=wb.active; ws.title="Main Market Data"
    write_sheet(ws,data,fmts)
    rev=matched[matched["Match Status"]!="High"].sort_values("Match Score")
    cols=[matched.columns[0],NAME_HEADER,"Ticker","Matched Company (CSV)","Match Score","Match Status","Closest (unconstrained)","Runner-up"]
    rev=rev[cols]
    ws2=wb.create_sheet("Match Review"); write_sheet(ws2,rev,{"Match Score":"0.0"})
    for i,r in enumerate(rev.to_dict("records"),2):
        ws2.cell(i,6).fill=PatternFill("solid",start_color=FILLS[r["Match Status"]])
    buf=BytesIO(); wb.save(buf); return buf.getvalue()

# ------------------------------ APP -------------------------------------------
pdf_file=st.file_uploader("Upload EGXReport.pdf",type=["pdf"])
csv_file=st.file_uploader("Upload egx_companies.csv",type=["csv"])

if csv_file is None and __import__("os").path.exists("egx_companies.csv"):
    st.info("Using local egx_companies.csv")
    csv_file="egx_companies.csv"

if st.button("Extract & Generate Excel",type="primary",disabled=pdf_file is None or csv_file is None):
    try:
        with st.spinner("Extracting PDF table and matching companies..."):
            raw=extract_table(BytesIO(pdf_file.getvalue()))
            companies=load_companies(csv_file)
            matched=match_tickers(raw,companies,NAME_HEADER,MANUAL_ALIASES)
            data,fmts,missing=select_export(matched,WANTED_COLUMNS)
            xlsx=excel_bytes(data,fmts,matched)
        st.session_state["result"]=(data,matched,xlsx,missing)
        st.success(f"Successfully extracted {len(raw):,} records.")
    except Exception as e:
        st.error(f"Extraction failed: {e}")

if "result" in st.session_state:
    data,matched,xlsx,missing=st.session_state["result"]
    st.subheader("Export Preview")
    st.dataframe(data,use_container_width=True,hide_index=True)
    st.download_button("Download Excel",xlsx,file_name="EGX_market_data.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",type="primary")
    st.subheader("Ticker Matching")
    st.write(matched["Match Status"].value_counts().to_dict())
    review=matched[matched["Match Status"]!="High"]
    if not review.empty:
        st.warning(f"{len(review)} matches require review.")
        st.dataframe(review[["Stock Name","Ticker","Matched Company (CSV)","Match Score","Match Status","Closest (unconstrained)","Runner-up"]],use_container_width=True,hide_index=True)
    if missing:
        st.warning("Missing PDF columns: "+", ".join(missing))
