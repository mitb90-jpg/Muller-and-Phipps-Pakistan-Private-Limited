"""
Bank Reconciliation Automation — Muller & Phipps Pakistan
Single file: engine + highlighting + Streamlit UI
Dealer code filters GL and each bank file.
Bank files assigned to bank names via dropdown in UI.
"""
import re
import io
import os
import zipfile
import tempfile
import shutil
import pandas as pd
import streamlit as st
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import CellIsRule

# ─────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Muller & Phipps Pakistan — Bank Reconciliation",
    page_icon="🏦", layout="wide", initial_sidebar_state="expanded"
)
st.markdown("""
<style>
.main-header{
    background:linear-gradient(135deg,#003366 0%,#00509e 100%);
    padding:20px 28px; border-radius:10px; color:white; margin-bottom:20px;
}
.company{font-size:30px;font-weight:bold;}
.system{font-size:18px;color:#d9e6f2;margin-top:4px;}
[data-testid="metric-container"]{
    background:#f0f4fa;border:1px solid #dce6f5;
    border-radius:10px;padding:14px 18px;
}
[data-testid="stSidebar"]{background:#f5f8ff;}
</style>
""", unsafe_allow_html=True)
st.markdown("""
<div class="main-header">
    <div class="company">Bank Reconciliation Automation System</div>
    <div class="system">🏦 Muller &amp; Phipps Pakistan (Private) Limited</div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# MASTER DATA
# ─────────────────────────────────────────────────────────────
DEALERS = {
    612: "Lahore Gulberg 3",
    613: "Lahore Hunjarwal Hub",
    614: "Lahore Usama Centre, Hall Road",
    645: "Shahpur Kaanjra CE",
    646: "Samsung CE Packages Mall",
    651: "Hassan Tower SC",
    657: "Zaman Plaza SC",
    733: "Techsirat Gulberg Lahore",
    766: "Techsirat Lahore Hunjarwal",
    778: "Techsirat Lahore Kot Lakhpat",
}

BANKS = {
    212201: "MCB Bank Ltd",
    212202: "Habib Bank Limited",
    212203: "National Bank of Pakistan",
    212204: "HBL Konnect",
    212205: "United Bank Limited",
    212206: "Bank Al Falah Limited",
    212207: "Bank Al Habib Limited",
    212208: "Samba Bank Limited",
    212209: "Faysal Bank Limited",
    212210: "Dubai Islamic Bank Limited",
    212211: "Meezan Bank Limited",
    212212: "Habib Metropoliton Bank",
    212213: "Meezan Bank Limited (MFS)",
    212214: "Habib Bank Limited (MFS)",
    212215: "Askari Bank Limited",
    212216: "Bank of Punjab",
    212217: "United Bank Limited (MFS)",
    212218: "Bank Islami Limited",
    212219: "UBL Omni",
    212503: "Standard Chartered Bank",
}

# Bank MIS column config — ds_col, chq_col, amt_col, mis_type
# mis_type: bulk=lump-sum confirms batch, line=per-cheque, direct=per-DS
BANK_CONFIG = {
    212201: dict(ds_col="Deposit Slip No.",  chq_col=None,        amt_col="Amount",  mis_type="bulk"),
    212202: dict(ds_col="Deposit Slip No.",  chq_col=None,        amt_col="Amount",  mis_type="bulk"),
    212205: dict(ds_col="Deposit Slip No.",  chq_col="CHEQUE NO", amt_col="AMOUNT",  mis_type="line"),
    212207: dict(ds_col="PAY SLIP NO",       chq_col="INST NO",   amt_col="AMOUNT",  mis_type="bulk"),
    212211: dict(ds_col="Deposit Slip",      chq_col=None,        amt_col="TOTAL",   mis_type="direct"),
}

TOL = 2.0

# ─────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────
def norm(x):
    if pd.isna(x) or x is None: return None
    s = re.sub(r"\.0$","",str(x).strip())
    s = re.sub(r"^[-]+","",s)
    s = re.sub(r"^0+(?=\d)","",s)
    return s or None

def extract_ds(text):
    if not isinstance(text,str): return None
    m = re.search(r"DS\s*#?\s*([A-Za-z0-9\-]+)", text, re.IGNORECASE)
    if m: return norm(m.group(1))
    if re.fullmatch(r"-?\d+",text.strip()): return norm(text.strip())
    return None

def is_lump(text):
    if not isinstance(text,str): return False
    t = text.lower()
    # must contain MIS + time reference BUT must NOT start with a date
    # (prior-period entries like "28-Apr-2026 HBL DS # ... FROM MIS ..." are NOT lump-sums)
    if re.match(r"\d{1,2}[-/][A-Za-z]{3}[-/]\d{4}", text.strip()):
        return False
    return "mis" in t and ("month" in t or "ftm" in t or "for the" in t)

def chq_suffix(a,b):
    if not a or not b: return False
    a,b = str(a),str(b)
    if a==b: return True
    lo,sh = (a,b) if len(a)>len(b) else (b,a)
    return len(sh)>=6 and lo.endswith(sh)

def gl_end_balance(sub):
    row = sub[["Period Ending Balance DR ","Period Ending Balance CR"]].dropna(how="all")
    if len(row):
        dr = pd.to_numeric(row["Period Ending Balance DR "].iloc[0],errors="coerce") or 0
        cr = pd.to_numeric(row["Period Ending Balance CR"].iloc[0], errors="coerce") or 0
        return dr-cr
    return 0

def load_bank_file(path, dealer_code, cfg):
    """Load bank file, filter by Dealer Code, return DS/CHQ/Amount df."""
    try:
        df = pd.read_excel(path, sheet_name=0)
        if "Dealer Code" in df.columns:
            df = df[df["Dealer Code"].astype(str).str.strip()==str(dealer_code)]
        if df.empty:
            return pd.DataFrame(columns=["DS","CHQ","Amount"])
        out = pd.DataFrame()
        out["DS"]     = df[cfg["ds_col"]].apply(norm) if cfg.get("ds_col") and cfg["ds_col"] in df.columns else None
        out["CHQ"]    = df[cfg["chq_col"]].apply(norm) if cfg.get("chq_col") and cfg["chq_col"] in df.columns else None
        out["Amount"] = pd.to_numeric(df[cfg["amt_col"]],errors="coerce").fillna(0) if cfg.get("amt_col") and cfg["amt_col"] in df.columns else 0
        out = out[out["DS"].notna() | out["CHQ"].notna()]
        out["DS"] = out.apply(lambda r: r["DS"] if pd.notna(r["DS"]) else r["CHQ"], axis=1)
        return out.reset_index(drop=True)
    except Exception:
        return pd.DataFrame(columns=["DS","CHQ","Amount"])

# ─────────────────────────────────────────────────────────────
# RECONCILIATION ENGINE
# ─────────────────────────────────────────────────────────────
def reconcile_bank(sub, bank_df, cfg, period_month):
    e1 = pd.DataFrame(columns=["Date","DS (Salesman)","DS (Bank MIS)","CHQ","Amount","Remark"])
    e4 = pd.DataFrame(columns=["Date","DS (Salesman)","DS (Bank MIS)","CHQ","Amount","Remark"])
    if sub.empty:
        return dict(annex1=e1,annex4=e4,k=0,lump=0,manual_flag=None)

    sal_dr = sub[(sub["GL Source"]=="Salezman")&sub["Entered Amount DR"].notna()].copy()
    sal_cr = sub[(sub["GL Source"]=="Salezman")&sub["Entered Amount CR"].notna()].copy()
    sp     = sub[sub["GL Source"]=="Spreadsheet"].copy()

    lump_total = sp[sp["Journal Line Description"].apply(is_lump)]["Entered Amount CR"].fillna(0).sum()
    indiv = sp[~sp["Journal Line Description"].apply(is_lump)&sp["Entered Amount CR"].notna()].copy()
    indiv["DS"] = indiv["Journal Line Description"].apply(extract_ds)

    period = pd.to_datetime(period_month)
    def is_prior(text):
        if not isinstance(text,str): return False
        m = re.match(r"(\d{1,2})[-/]([A-Za-z]{3})[-/](\d{4})",text.strip())
        if not m: return False
        mm={"jan":1,"feb":2,"mar":3,"apr":4,"may":5,"jun":6,
            "jul":7,"aug":8,"sep":9,"oct":10,"nov":11,"dec":12}
        try:
            ey,em=int(m.group(3)),mm[m.group(2).lower()]
            return (ey,em)<(period.year,period.month)
        except: return False

    current_cr = indiv[~indiv["Journal Line Description"].apply(is_prior)]
    k   = gl_end_balance(sub)
    mis = cfg.get("mis_type","bulk")

    # ── ANNEX 1 ──────────────────────────────────────────────
    rows1=[]
    sal_dr_total = sal_dr["Entered Amount DR"].sum()

    if lump_total >= sal_dr_total-TOL:
        pass  # full lump-sum — all knocked off

    elif mis=="bulk" and lump_total>0 and abs(k)>TOL:
        blist=list(bank_df[["DS","CHQ","Amount"]].itertuples(index=False))
        used=set()
        for _,r in sal_dr.iterrows():
            ref=norm(str(r["Journal Line Description"]))
            ds=norm(r.get("DS_extracted"))
            amt=r["Entered Amount DR"]
            hit=False
            for i,b in enumerate(blist):
                if i in used: continue
                if (b.CHQ and b.CHQ==ref) or (b.DS and (b.DS==ref or b.DS==ds)):
                    if abs(b.Amount-amt)<=TOL:
                        used.add(i); hit=True; break
            if not hit:
                rows1.append({"Date":r["GL Date"],"DS (Salesman)":ds or ref,
                               "DS (Bank MIS)":None,"CHQ":ref,"Amount":amt,
                               "Remark":"Salesman deposit not yet confirmed by bank"})

    elif mis=="direct":
        cr_amts=sorted(current_cr["Entered Amount CR"].tolist())
        for _,r in sal_dr.sort_values("Entered Amount DR").iterrows():
            amt=r["Entered Amount DR"]
            idx=next((i for i,ca in enumerate(cr_amts) if abs(ca-amt)<=TOL),None)
            if idx is not None: cr_amts.pop(idx)
            else:
                rows1.append({"Date":r["GL Date"],
                               "DS (Salesman)":norm(str(r["Journal Line Description"])),
                               "DS (Bank MIS)":None,"CHQ":None,"Amount":amt,
                               "Remark":"Salesman deposit not yet confirmed by bank"})

    annex1=pd.DataFrame(rows1) if rows1 else e1

    # ── ANNEX 4 ──────────────────────────────────────────────
    rows4=[]

    if mis=="line":
        sal_map={norm(str(r["Journal Line Description"])):r["Entered Amount DR"]
                 for _,r in sal_dr.iterrows() if norm(str(r["Journal Line Description"]))}
        for _,b in bank_df.iterrows():
            bchq=norm(b["CHQ"]); bamt=b["Amount"]; bds=norm(b["DS"])
            sr=next((s for s in sal_map if chq_suffix(bchq,s)),None)
            if sr is None:
                rows4.append({"Date":None,"DS (Salesman)":None,"DS (Bank MIS)":bds,
                               "CHQ":bchq,"Amount":bamt,
                               "Remark":"*** NEEDS REVIEW: Bank credit not in GL - check if cross-GL/branch"})
            else:
                excess=round(bamt-sal_map[sr],2)
                if excess>0.01:
                    rows4.append({"Date":None,"DS (Salesman)":sr,"DS (Bank MIS)":bds,
                                   "CHQ":bchq,"Amount":excess,"Remark":"EXTRA CREDIT (rounding)"})

    elif mis=="direct":
        bank_totals=bank_df.groupby("DS")["Amount"].sum().to_dict() if len(bank_df) else {}
        for _,r in current_cr.iterrows():
            ds=norm(r.get("DS_extracted") or r.get("DS"))
            amt=r["Entered Amount CR"]
            if abs(bank_totals.get(ds,0)-amt)>TOL:
                rows4.append({"Date":r["GL Date"],"DS (Salesman)":ds,"DS (Bank MIS)":ds,
                               "CHQ":None,"Amount":amt,
                               "Remark":"Collection posted in GL, not yet in Bank MIS (timing difference)"})

    else:  # bulk
        bank_ds = set(bank_df["DS"].dropna()) if len(bank_df) else set()
        sal_amts = set(round(float(x), 0) for x in sal_dr["Entered Amount DR"]) if len(sal_dr) else set()
        for _, r in current_cr.iterrows():
            ds  = norm(r.get("DS_extracted") or r.get("DS"))
            amt = r["Entered Amount CR"]
            amt_r = round(float(amt), 0)
            # skip if this CR amount knocks off against a salesman DR amount
            # (means it's already reconciled internally)
            if amt_r in sal_amts:
                continue
            # skip if confirmed in bank MIS
            if ds in bank_ds:
                continue
            rows4.append({"Date":r["GL Date"],"DS (Salesman)":ds,"DS (Bank MIS)":ds,
                           "CHQ":None,"Amount":amt,
                           "Remark":"Collection posted in GL, not yet in Bank MIS (timing difference)"})

    annex4=pd.DataFrame(rows4) if rows4 else e4
    a1t=annex1["Amount"].sum() if len(annex1) else 0
    a4t=annex4["Amount"].sum() if len(annex4) else 0
    j=a1t-a4t
    flag=None
    if abs(k-j)>TOL:
        flag=f"Gap of Rs {k-j:,.2f} — cross-GL or cross-branch items need manual input."
    return dict(annex1=annex1,annex4=annex4,k=k,lump=lump_total,manual_flag=flag)


def run_reconciliation(gl_path, bank_assignments, dealer_code, month):
    """
    gl_path         : combined GL file path (sheet=GL, col=Dealer Code)
    bank_assignments: dict {gl_code(int): file_path} — user assigned in UI
    dealer_code     : int
    month           : str YYYY-MM-DD
    """
    gl_df = pd.read_excel(gl_path, sheet_name="GL")
    gl_df["DS_extracted"] = gl_df["Journal Line Description"].apply(extract_ds)
    # filter GL by dealer code
    dealer_gl = gl_df[gl_df["Dealer Code"].astype(str).str.strip()==str(dealer_code)].copy()

    results={}
    for gl_code, bank_name in BANKS.items():
        sub = dealer_gl[dealer_gl["Natural Account Segment"]==gl_code].copy()
        cfg = BANK_CONFIG.get(gl_code, {})
        bank_df = pd.DataFrame(columns=["DS","CHQ","Amount"])
        if gl_code in bank_assignments and cfg:
            bank_df = load_bank_file(bank_assignments[gl_code], dealer_code, cfg)
        r = reconcile_bank(sub, bank_df, cfg, month)
        if r["k"]!=0 or len(r["annex1"]) or len(r["annex4"]):
            results[bank_name] = dict(cfg=dict(cfg, gl_code=gl_code), **r)

    return results, dealer_gl

# ─────────────────────────────────────────────────────────────
# TEMPLATE POPULATION
# ─────────────────────────────────────────────────────────────
YELLOW = PatternFill("solid", fgColor="FFE699")
RED_CF = PatternFill("solid", fgColor="FFC7CE")

def find_blocks(ws):
    blocks=[]
    for row in range(1,ws.max_row+1):
        v=ws.cell(row=row,column=1).value
        if isinstance(v,str) and "Main Reconciliation" in v and "CONCATENATE" in v:
            m=re.search(r"!B(\d+)",v)
            if m: blocks.append((row,int(m.group(1))))
    return blocks

def fill_annex(ws, main_row, items, col_map):
    blocks=find_blocks(ws)
    hrow=next(r for r,mr in blocks if mr==main_row)
    dstart=hrow+2
    r=dstart
    while r<dstart+60:
        v=ws.cell(row=r,column=7).value
        if isinstance(v,str) and "SUM" in v: break
        r+=1
    sum_row=r; capacity=sum_row-dstart
    n=len(items)
    if n>capacity:
        ws.insert_rows(sum_row,amount=n-capacity)
        sum_row+=(n-capacity); capacity=n
    for i in range(capacity):
        row=dstart+i
        ws.cell(row=row,column=1,value=i+1)
        if i<n:
            it=items[i]
            for col_idx,key in col_map.items():
                ws.cell(row=row,column=col_idx,value=it.get(key))
            ws.cell(row=row,column=7,value=round(it["Amount"],2))
            ws.cell(row=row,column=9,value=it.get("Remark",""))
            for c in range(1,10): ws.cell(row=row,column=c).fill=YELLOW
        else:
            ws.cell(row=row,column=7,value=0)
    ws.cell(row=sum_row,column=7,value=f"=SUM(G{dstart}:G{sum_row-1})")
    return sum_row

def populate_workbook(template_path, results, dealer_code, dealer_name, month, out_path):
    wb=load_workbook(template_path)
    main=wb["Main Reconciliation"]
    ws1=wb["Annex 1 (Debit M&P)"]
    ws4=wb["Annex 4 (Credit Bank)"]
    blocks1=find_blocks(ws1); blocks4=find_blocks(ws4)
    name_to_mr1={main.cell(row=mr,column=2).value:mr for _,mr in blocks1}
    name_to_mr4={main.cell(row=mr,column=2).value:mr for _,mr in blocks4}
    col_map={2:"Date",3:"DS (Salesman)",4:"DS (Bank MIS)",6:"CHQ"}
    for name,mr in sorted(name_to_mr1.items(),key=lambda x:x[1]):
        r=results.get(name)
        items=r["annex1"].to_dict("records") if r and len(r["annex1"]) else []
        tr=fill_annex(ws1,mr,items,col_map)
        main.cell(row=mr,column=6,value=f"=+'Annex 1 (Debit M&P)'!G{tr}")
    for name,mr in sorted(name_to_mr4.items(),key=lambda x:x[1]):
        r=results.get(name)
        items=r["annex4"].to_dict("records") if r and len(r["annex4"]) else []
        tr=fill_annex(ws4,mr,items,col_map)
        main.cell(row=mr,column=9,value=f"=-'Annex 4 (Credit Bank)'!G{tr}")
    for name,mr in sorted(name_to_mr1.items(),key=lambda x:x[1]):
        r=results.get(name)
        main.cell(row=mr,column=3,value=int(dealer_code))
        if r: main.cell(row=mr,column=11,value=round(r["k"],2))
    main["L4"]=month
    first_mr=min(name_to_mr1.values()); last_mr=max(name_to_mr1.values())
    main.conditional_formatting.add(f"L{first_mr}:L{last_mr}",
        CellIsRule(operator="notBetween",formula=["-1","1"],fill=RED_CF))
    wb.calculation.calcMode="auto"
    wb.save(out_path)

# ─────────────────────────────────────────────────────────────
# HIGHLIGHTING
# ─────────────────────────────────────────────────────────────
B_FILL=PatternFill("solid",fgColor="BDD7EE")
G_FILL=PatternFill("solid",fgColor="C6EFCE")
R_FILL=PatternFill("solid",fgColor="FFC7CE")

def highlight_gl(gl_path, dealer_gl, results, out_path):
    shutil.copy(gl_path,out_path)
    wb=load_workbook(out_path); ws=wb.active; max_c=ws.max_column
    row_colour={}
    for name,res in results.items():
        gl_code=res["cfg"].get("gl_code")
        sub=dealer_gl[dealer_gl["Natural Account Segment"]==gl_code]
        a1_amts=set(round(float(x),0) for x in res["annex1"]["Amount"]) if len(res["annex1"]) else set()
        a4_amts=set(round(float(x),0) for x in res["annex4"]["Amount"]) if len(res["annex4"]) else set()
        has_exc=bool(res.get("manual_flag"))
        for idx,row in sub.iterrows():
            dr=row.get("Entered Amount DR"); cr=row.get("Entered Amount CR")
            amt=dr if pd.notna(dr) else cr
            if pd.isna(amt): continue
            amt_r=round(float(amt),0)
            desc=str(row.get("Journal Line Description",""))
            src=str(row.get("GL Source",""))
            if is_lump(desc): row_colour[idx]=B_FILL
            elif amt_r in a1_amts or amt_r in a4_amts: row_colour[idx]=G_FILL
            elif has_exc and src=="Spreadsheet" and not is_lump(desc): row_colour[idx]=R_FILL
            else: row_colour[idx]=B_FILL
    for excel_row in range(2,ws.max_row+1):
        fill=row_colour.get(excel_row-2)
        if fill:
            for c in range(1,max_c+1): ws.cell(row=excel_row,column=c).fill=fill
    wb.save(out_path)

def highlight_bank_files(bank_assignments, dealer_code, results, out_dir):
    highlighted={}
    for gl_code, fpath in bank_assignments.items():
        bank_name=BANKS.get(gl_code,"")
        res=results.get(bank_name,{})
        cfg=BANK_CONFIG.get(gl_code,{})
        if not cfg.get("amt_col"): continue
        try:
            a4_amts=set(round(float(x),0) for x in res["annex4"]["Amount"]) if res.get("annex4") is not None and len(res["annex4"]) else set()
            has_exc=bool(res.get("manual_flag")) if res else False
            out_path=os.path.join(out_dir,f"Bank_Highlighted_{bank_name.replace(' ','_')}.xlsx")
            shutil.copy(fpath,out_path)
            wb=load_workbook(out_path); ws=wb.active; max_c=ws.max_column
            headers={ws.cell(row=1,column=c).value:c for c in range(1,max_c+1)}
            amt_c=headers.get(cfg["amt_col"]); dep_c=headers.get("Dealer Code")
            if not amt_c: continue
            for r in range(2,ws.max_row+1):
                if dep_c:
                    dep_val=str(ws.cell(row=r,column=dep_c).value or "").strip()
                    if dep_val!=str(dealer_code): continue
                amt_v=ws.cell(row=r,column=amt_c).value
                if amt_v is None: continue
                try: amt_r=round(float(amt_v),0)
                except: continue
                fill=G_FILL if amt_r in a4_amts else (R_FILL if has_exc else B_FILL)
                for c in range(1,max_c+1): ws.cell(row=r,column=c).fill=fill
            wb.save(out_path)
            highlighted[bank_name]=out_path
        except Exception:
            continue
    return highlighted

# ─────────────────────────────────────────────────────────────
# SUMMARY
# ─────────────────────────────────────────────────────────────
def build_summary(results):
    rows=[]
    for name,r in results.items():
        a1=r["annex1"]["Amount"].sum() if len(r["annex1"]) else 0
        a4=r["annex4"]["Amount"].sum() if len(r["annex4"]) else 0
        k=r["k"]; j=a1-a4; diff=round(k-j,2)
        rows.append({
            "Bank":name,
            "Annex 1 (A1)":round(a1,2),
            "Annex 4 (A4)":round(a4,2),
            "GL Bal per Recon (J)":round(j,2),
            "Actual GL Bal (K)":round(k,2),
            "Difference (K-J)":diff,
            "Status":"✅ Reconciled" if abs(diff)<=1 else "⚠️ Needs Review",
            "Flag":r.get("manual_flag") or "",
        })
    return pd.DataFrame(rows)

# ─────────────────────────────────────────────────────────────
# ─────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### Navigation")
    page=st.radio("",[
        "📊 Dashboard",
        "🔄 Run Reconciliation",
        "📋 Review Results",
        "📥 Download Output"
    ],label_visibility="collapsed")
    st.markdown("---")
    st.markdown("### Period Settings")
    dealer_options={f"{code} — {name}":code for code,name in DEALERS.items()}
    selected_label=st.selectbox("Select Dealer",list(dealer_options.keys()))
    dealer_code=dealer_options[selected_label]
    dealer_name=DEALERS[dealer_code]
    month=st.date_input("Month-End Date",value=datetime(2026,5,31))
    month_str=month.strftime("%Y-%m-%d")
    st.markdown("---")
    st.markdown("### Upload Files")
    gl_file=st.file_uploader("Combined GL File (.xlsx)",type="xlsx",key="gl_upload")
    bank_uploads=st.file_uploader("Bank MIS Files (.xlsx)",type="xlsx",
                                   accept_multiple_files=True,key="bank_upload")
    tpl_file=st.file_uploader("Last Month Template (.xlsx)",type="xlsx",key="tpl_upload")
    if tpl_file:
        st.session_state["tpl_bytes"]=tpl_file.read()

# ─────────────────────────────────────────────────────────────
# ─────────────────────────────────────────────────────────────
# RUN RECONCILIATION PAGE
# ─────────────────────────────────────────────────────────────
if page=="🔄 Run Reconciliation":
    st.markdown("## Run Reconciliation")
    st.markdown("""
**Monthly prep before uploading:**
| File | What to rename |
|---|---|
| GL file | Sheet → `GL`, dealer code column → `Dealer Code` |
| Each bank file | Dealer code column → `Dealer Code` |
""")
    st.markdown("---")
    st.markdown("### Assign Bank Files")
    st.caption("Upload each bank's MIS file in the sidebar, then assign each file to the correct bank below.")

    # Auto-match keywords for each bank
    BANK_KEYWORDS = {
        212201: ["mcb"],
        212202: ["hbl", "habib bank"],
        212203: ["nbl", "national bank"],
        212204: ["hbl konnect", "konnect"],
        212205: ["ubl", "united bank"],
        212206: ["falah", "baf", "alfalah"],
        212207: ["bahl", "al habib", "alhabib"],
        212208: ["samba"],
        212209: ["faysal"],
        212210: ["dubai", "dib"],
        212211: ["meezan"],
        212212: ["metro", "hmb"],
        212213: ["meezan mfs"],
        212214: ["hbl mfs"],
        212215: ["askari"],
        212216: ["bop", "punjab"],
        212217: ["ubl mfs"],
        212218: ["islami"],
        212219: ["omni"],
        212503: ["standard", "scb"],
    }

    def auto_detect_bank(filename):
        name = filename.lower().replace("_"," ").replace("-"," ")
        for gl_code, keywords in BANK_KEYWORDS.items():
            if any(kw in name for kw in keywords):
                return BANKS[gl_code]
        return "— Not uploaded —"

    bank_assignments={}
    if bank_uploads:
        bank_names_list=["— Not uploaded —"]+list(BANKS.values())
        st.markdown("**Assign each uploaded file to its bank (auto-detected where possible):**")
        for bf in bank_uploads:
            auto = auto_detect_bank(bf.name)
            col_a,col_b=st.columns([2,3])
            with col_a:
                st.markdown(f"📄 `{bf.name}`")
            with col_b:
                chosen=st.selectbox(
                    f"Bank for {bf.name}",
                    bank_names_list,
                    index=bank_names_list.index(auto) if auto in bank_names_list else 0,
                    key=f"assign_{bf.name}",
                    label_visibility="collapsed"
                )
                )
            if chosen!="— Not uploaded —":
                gl_code=next((c for c,n in BANKS.items() if n==chosen),None)
                if gl_code:
                    with tempfile.NamedTemporaryFile(suffix=".xlsx",delete=False) as tmp:
                        tmp.write(bf.getbuffer())
                        bank_assignments[gl_code]=tmp.name

    if bank_assignments:
        st.session_state["bank_assignments"]=bank_assignments

    bank_assignments=bank_assignments or st.session_state.get("bank_assignments",{})

    st.markdown("---")
    run_btn=st.button("▶  Run Reconciliation",type="primary",
                      disabled=not(gl_file),
                      use_container_width=True)
    if not gl_file:
        st.caption("Upload GL file in the sidebar to enable.")

    if run_btn and gl_file:
        with tempfile.NamedTemporaryFile(suffix=".xlsx",delete=False) as fg:
            fg.write(gl_file.getbuffer()); gl_path=fg.name
        try:
            with st.spinner(f"Reconciling {dealer_code} — {dealer_name}…"):
                results,dealer_gl=run_reconciliation(gl_path,bank_assignments,dealer_code,month_str)
            gl_hl=tempfile.NamedTemporaryFile(suffix=".xlsx",delete=False).name
            highlight_gl(gl_path,dealer_gl,results,gl_hl)
            tmp_dir=tempfile.mkdtemp()
            bank_hl_map=highlight_bank_files(bank_assignments,dealer_code,results,tmp_dir)
            st.session_state.update({
                "results":results,"dealer_gl":dealer_gl,
                "month_str":month_str,"dealer_code":dealer_code,"dealer_name":dealer_name,
                "gl_path":gl_path,"bank_assignments":bank_assignments,
                "gl_hl":gl_hl,"bank_hl_map":bank_hl_map,
            })
            st.success(f"✅ Done — Dealer {dealer_code} | {dealer_name} | {month_str}")
        except Exception as e:
            st.error(f"Engine error: {e}")

# ─────────────────────────────────────────────────────────────
# OTHER PAGES
# ─────────────────────────────────────────────────────────────
results    =st.session_state.get("results",{})
month_str  =st.session_state.get("month_str",month_str)
dealer_code=st.session_state.get("dealer_code",dealer_code)
dealer_name=st.session_state.get("dealer_name",dealer_name)

if page=="📊 Dashboard":
    st.markdown("## Dashboard")
    if not results:
        st.info("Go to **🔄 Run Reconciliation** to upload files and run.")
    else:
        df=build_summary(results)
        n_ok=int((df["Status"].str.contains("✅")).sum())
        n_warn=len(df)-n_ok
        exc=int((df["Annex 1 (A1)"]!=0).sum()+(df["Annex 4 (A4)"]!=0).sum())
        rate=f"{round(n_ok/len(df)*100,2)}%" if len(df) else "—"
        k1,k2,k3,k4=st.columns(4)
        k1.metric("🏦 Bank Accounts",len(df))
        k2.metric("📂 Open Reconciliations",n_warn)
        k3.metric("⚠️ Pending Exceptions",exc)
        k4.metric("✅ Match Rate",rate)
        st.write("")
        left,right=st.columns([2,1])
        with left:
            st.subheader("Reconciliation Status")
            def _s(row):
                bg="#c6efce" if "✅" in row["Status"] else "#ffeb9c"
                return [f"background-color:{bg}"]*len(row)
            cols=["Bank","GL Bal per Recon (J)","Actual GL Bal (K)","Difference (K-J)","Status"]
            st.dataframe(df[cols].style.apply(_s,axis=1),use_container_width=True,hide_index=True)
        with right:
            st.subheader("Quick Actions")
            st.markdown(f"**Dealer:** `{dealer_code}` — {dealer_name}")
            st.markdown(f"**Period:** `{month_str}`")
            st.write("")
            if n_warn: st.warning(f"{n_warn} bank(s) need review → **Review Results**")
            else: st.success("All banks reconciled ✅ → **Download Output**")
        st.write("")
        left2,right2=st.columns(2)
        with left2:
            st.subheader("Recent Activity")
            for _,row in df.iterrows():
                icon="✅" if "✅" in row["Status"] else "⚠️"
                st.markdown(f"{icon} **{row['Bank']}** — Diff: Rs {row['Difference (K-J)']:,.2f}")
        with right2:
            st.subheader("Pending Exceptions")
            flagged=df[df["Flag"]!=""]
            if len(flagged):
                for _,row in flagged.iterrows():
                    st.warning(f"**{row['Bank']}:** {row['Flag'][:120]}")
            else:
                st.success("No exceptions requiring manual input.")

elif page=="📋 Review Results":
    st.markdown("## Review Results")
    if not results:
        st.info("Run the reconciliation first.")
    else:
        df=build_summary(results)
        def _s(row):
            bg="#c6efce" if "✅" in row["Status"] else "#ffeb9c"
            return [f"background-color:{bg}"]*len(row)
        st.dataframe(df.drop(columns=["Flag"]).style.apply(_s,axis=1),
                     use_container_width=True,hide_index=True)
        st.write("")
        active=list(results.keys())
        if active:
            sel=st.selectbox("Drill into a bank:",active)
            r=results[sel]
            c1,c2=st.columns(2)
            with c1:
                st.markdown("**Annex 1 — Debit in M&P Ledger, not yet in Bank**")
                if len(r["annex1"]): st.dataframe(r["annex1"],use_container_width=True,hide_index=True)
                else: st.success("Nothing in Annex 1.")
            with c2:
                st.markdown("**Annex 4 — Credit in Bank, not yet in GL**")
                if len(r["annex4"]): st.dataframe(r["annex4"],use_container_width=True,hide_index=True)
                else: st.success("Nothing in Annex 4.")
            if r.get("manual_flag"):
                st.warning(f"⚠️ {r['manual_flag']}")

elif page=="📥 Download Output":
    st.markdown("## Download Output")
    if not results:
        st.info("Run the reconciliation first.")
    else:
        tpl_bytes=st.session_state.get("tpl_bytes")
        if not tpl_bytes:
            st.warning("Upload the **Last Month Template** on the Run Reconciliation page.")
        else:
            if st.button("📄 Generate Populated Workbook",type="primary"):
                with tempfile.NamedTemporaryFile(suffix=".xlsx",delete=False) as ft:
                    ft.write(tpl_bytes); tpl_path=ft.name
                out_path=tpl_path.replace(".xlsx","_out.xlsx")
                try:
                    populate_workbook(tpl_path,results,dealer_code,dealer_name,month_str,out_path)
                    with open(out_path,"rb") as f: out_bytes=f.read()
                    fname=f"{dealer_code}_Reconciliation_DRAFT_{month.strftime('%b%Y')}.xlsx"
                    st.download_button("⬇️ Download Reconciliation File",out_bytes,
                        file_name=fname,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    st.success(f"Open **{fname}** in Excel — recalculates automatically.")
                except Exception as e:
                    st.error(f"Error: {e}")

        st.markdown("---")
        st.markdown("### 🎨 Highlighted Source Files")
        st.markdown("🔵 Blue = knocked off &nbsp;|&nbsp; 🟢 Green = in annexure &nbsp;|&nbsp; 🔴 Red = needs manual input")
        gl_hl=st.session_state.get("gl_hl")
        bank_hl_map=st.session_state.get("bank_hl_map",{})
        if gl_hl:
            with open(gl_hl,"rb") as f:
                st.download_button("⬇️ Download Highlighted GL",f.read(),
                    file_name=f"GL_Highlighted_{dealer_code}_{month_str}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        if bank_hl_map:
            zip_buf=io.BytesIO()
            with zipfile.ZipFile(zip_buf,"w") as zf:
                for bname,bpath in bank_hl_map.items():
                    zf.write(bpath,arcname=f"Bank_Highlighted_{bname.replace(' ','_')}.xlsx")
            st.download_button("⬇️ Download Highlighted Bank Files (ZIP)",zip_buf.getvalue(),
                file_name=f"Bank_Highlighted_{dealer_code}_{month_str}.zip",
                mime="application/zip")

        st.markdown("---")
        st.markdown("### Export Summary")
        df=build_summary(results)
        buf=io.BytesIO()
        df.to_excel(buf,index=False)
        st.download_button("⬇️ Download Summary (.xlsx)",buf.getvalue(),
            file_name=f"Recon_Summary_{dealer_code}_{month_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
