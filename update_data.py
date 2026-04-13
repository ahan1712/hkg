#!/usr/bin/env python3
"""
HKG Dashboard - Daily Data Updater
Run this script every day after updating your Excel file.
It reads your Excel and generates data.json for the dashboard.

Usage:
  python3 update_data.py
  python3 update_data.py /path/to/your/excel.xlsx
"""

import json, sys, os
from datetime import datetime
import numpy as np

try:
    import pandas as pd
    from openpyxl import load_workbook
except ImportError:
    print("Installing required packages...")
    os.system("pip3 install pandas openpyxl --break-system-packages -q")
    import pandas as pd

# ── CONFIG ─────────────────────────────────────────────────────────────────
# Put the name of your Excel file here (or pass as argument)
EXCEL_FILE = "all_inputs_final.xlsx"
OUTPUT_FILE = "data.json"

if len(sys.argv) > 1:
    EXCEL_FILE = sys.argv[1]

if not os.path.exists(EXCEL_FILE):
    print(f"ERROR: Could not find {EXCEL_FILE}")
    print("Usage: python3 update_data.py your_excel_file.xlsx")
    sys.exit(1)

print(f"Reading: {EXCEL_FILE}")

# ── LOAD SHEETS ────────────────────────────────────────────────────────────
def load(sheet, hrow):
    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet, header=hrow)
    df.columns = [str(c).replace("\n"," ").strip() for c in df.columns]
    df["Date"] = pd.to_datetime(df.iloc[:,0], errors="coerce")
    return df[df["Date"].notna() & df["Date"].dt.year.between(2024,2030)].reset_index(drop=True)

print("Loading sheets...", end=" ")
oci   = load("OtherCostInput", 2)
scrap = load("ScrapInput", 1)
sales = load("SalesInput", 1)
deliv = load("DeliveryInput", 1)
comp  = load("CompetitorInput", 1)
inv   = load("InventoryInput", 2)
orev  = load("OtherRevenueInput", 1)
print("done")

# ── LATEST DATE ────────────────────────────────────────────────────────────
latest = oci[oci["Rod Production Total (MT)"].notna() & (oci["Rod Production Total (MT)"]>0)]["Date"].max()
mtd_start = latest.replace(day=1)
print(f"Latest production date: {latest.date()}")
print(f"MTD start: {mtd_start.date()}")

# ── HELPERS ────────────────────────────────────────────────────────────────
def safe(v, dec=0):
    try:
        f = float(v)
        if f != f: return 0  # NaN check
        return round(f, dec) if dec > 0 else int(round(f))
    except: return 0

def get_period(days=None, mstart=None):
    end = latest
    def sl(df):
        if mstart: return df[(df["Date"]>=mstart)&(df["Date"]<=end)]
        elif days==0: return df[df["Date"]==end]
        else: return df[(df["Date"]>end-pd.Timedelta(days=days))&(df["Date"]<=end)]
    o,sa,dl,sc,or_ = sl(oci),sl(sales),sl(deliv),sl(scrap),sl(orev)
    rod = safe(o["Rod Production Total (MT)"].sum(), 2)
    sc_cost = safe(o["Total Scrap Consumption Cost (৳)"].sum())
    oc = safe(o["Total Other Cost Expenditure (৳)"].sum())
    sv = safe(sa["Total Order Size (MT)"].sum(), 2)
    rev = safe(sa["Total Sale Value (৳)"].sum())
    orv = safe(or_["Total Sale Value (৳)"].sum())
    tot_rev = rev + orv
    hms_e = safe((sc["HMS Qty Purchased (MT)"]*sc["HMS Price / Ton (৳)"]).sum())
    bun_e = safe((sc["Bundle Qty Purchased (MT)"]*sc["Bundle Price / Ton (৳)"]).sum())
    qp = safe(sc["HMS Qty Purchased (MT)"].sum() + sc["Bundle Qty Purchased (MT)"].sum(), 2)
    sc_cons = safe(o["Scrap Consumed HMS (MT)"].fillna(0).sum() + o["Scrap Consumed Bundle (MT)"].fillna(0).sum(), 2)
    billet = safe(o["Billet Produced (MT)"].fillna(0).sum(), 2)
    del_mt = safe(dl["Quantity of Rod (MT)"].sum(), 2)
    scrap_exp_total = safe(sc["Daily Scrap Expenditure (৳)"].sum())
    return {
        "rod": rod, "sales_vol": sv, "revenue": rev, "other_rev": orv, "total_rev": tot_rev,
        "delivery": del_mt, "trucks": len(dl),
        "electricity": safe(o["Electricity"].sum()),
        "avg_price": safe(rev/sv) if sv>0 else 0,
        "scrap_exp_pt": safe((hms_e+bun_e)/qp) if qp>0 else 0,
        "scrap_consump_pt": safe(sc_cost/rod) if rod>0 else 0,
        "other_cost_pt": safe(oc/rod) if rod>0 else 0,
        "total_cost_pt": safe((sc_cost+oc)/rod) if rod>0 else 0,
        "gross_margin": tot_rev-sc_cost-oc,
        "margin_pt": safe((tot_rev-sc_cost-oc)/rod) if rod>0 else 0,
        "sc_consumed": sc_cons, "billet": billet,
        "scrap_exp_total": scrap_exp_total, "oc_total": oc,
        "pct_delivered": round(del_mt/sv*100,1) if sv>0 else 0,
        "scrap_to_billet": round(sc_cons/billet,3) if billet>0 else 0,
    }

print("Computing periods...", end=" ")
periods = {
    "today": get_period(days=0),
    "mtd":   get_period(mstart=mtd_start),
    "7d":    get_period(days=7),
    "10d":   get_period(days=10),
    "14d":   get_period(days=14),
    "30d":   get_period(days=30),
}
print("done")

# ── COMPETITORS ────────────────────────────────────────────────────────────
comp_cols = ["Salam Steel Dealer Rate (৳)","RICL Dealer Rate (৳)",
             "JSRM Dealer Rate (৳)","Rani Dealer Rate (৳)","SAS Dealer Rate (৳)"]
cl = comp[comp["Date"]<=latest].sort_values("Date").tail(1).iloc[0]
others_avg = int(round(float(np.mean([float(cl[c]) for c in comp_cols if pd.notna(cl[c]) and float(cl[c])>0]))))
bsrm = int(float(cl["BSRM Dealer Rate (৳)"]))
hkg = periods["today"]["avg_price"]

# ── INVENTORY STOCKS ───────────────────────────────────────────────────────
il = inv[inv["Date"]<=latest].sort_values("Date").tail(1).iloc[0]
stocks = {
    "hms":    round(max(safe(il["Closing (Calc)"], 3), 0), 3),
    "bundle": round(max(safe(il["Closing (Calc).1"], 3), 0), 3),
    "billet": round(max(safe(il["Closing (Calc).2"], 3), 0), 3),
    "rod":    round(max(safe(il["Closing (Calc).3"], 3), 0), 3),
}

# ── TOP CUSTOMERS MTD ──────────────────────────────────────────────────────
cust_mtd = sales[(sales["Date"]>=mtd_start)&(sales["Date"]<=latest)]
cmap = {}
for _, r in cust_mtd.iterrows():
    nm = str(r.get("Customer Name",""))
    if not nm: continue
    if nm not in cmap: cmap[nm] = {"name":nm,"vol":0,"rev":0}
    cmap[nm]["vol"] += safe(r["Total Order Size (MT)"],2)
    cmap[nm]["rev"] += safe(r["Total Sale Value (৳)"])
top_custs = sorted(cmap.values(), key=lambda x:-x["rev"])[:8]
for c in top_custs: c["avg_px"] = safe(c["rev"]/c["vol"]) if c["vol"]>0 else 0

# ── SALESMEN MTD ───────────────────────────────────────────────────────────
smap = {}
for _, r in cust_mtd.iterrows():
    nm = str(r.get("Salesman Name",""))
    if not nm or nm=="Factory": continue
    if nm not in smap: smap[nm] = {"name":nm,"vol":0,"rev":0}
    smap[nm]["vol"] += safe(r["Total Order Size (MT)"],2)
    smap[nm]["rev"] += safe(r["Total Sale Value (৳)"])
salesmen = sorted(smap.values(), key=lambda x:-x["rev"])

# ── MONTHLY ────────────────────────────────────────────────────────────────
monthly = []
for prd, grp in oci.groupby(oci["Date"].dt.to_period("M")):
    ms=grp["Date"].min(); me=grp["Date"].max()
    s_=sales[(sales["Date"]>=ms)&(sales["Date"]<=me)]
    d_=deliv[(deliv["Date"]>=ms)&(deliv["Date"]<=me)]
    or_=orev[(orev["Date"]>=ms)&(orev["Date"]<=me)]
    rod=safe(grp["Rod Production Total (MT)"].sum(),2)
    if rod==0: continue
    sc=safe(grp["Total Scrap Consumption Cost (৳)"].sum())
    oc=safe(grp["Total Other Cost Expenditure (৳)"].sum())
    sv=safe(s_["Total Order Size (MT)"].sum(),2)
    rev=safe(s_["Total Sale Value (৳)"].sum())
    orv=safe(or_["Total Sale Value (৳)"].sum())
    tr=rev+orv
    monthly.append({
        "month":str(prd), "label":ms.strftime("%b %Y"),
        "rod":rod, "sales_vol":sv, "revenue":tr,
        "avg_price":safe(rev/sv) if sv>0 else 0,
        "total_cost_pt":safe((sc+oc)/rod) if rod>0 else 0,
        "scrap_consump_pt":safe(sc/rod) if rod>0 else 0,
        "other_cost_pt":safe(oc/rod) if rod>0 else 0,
        "margin_pt":safe((tr-sc-oc)/rod) if rod>0 else 0,
        "gross_margin":tr-sc-oc,
        "electricity":safe(grp["Electricity"].sum()),
        "scrap_consumed":safe(grp["Scrap Consumed HMS (MT)"].fillna(0).sum()+grp["Scrap Consumed Bundle (MT)"].fillna(0).sum(),2),
        "delivery":safe(d_["Quantity of Rod (MT)"].sum(),2),
    })

# ── WRITE JSON ─────────────────────────────────────────────────────────────
data = {
    "meta": {
        "latest_date": latest.strftime("%d %b %Y"),
        "latest_date_iso": latest.strftime("%Y-%m-%d"),
        "mtd_start": mtd_start.strftime("%d %b %Y"),
        "generated": datetime.now().strftime("%d %b %Y %H:%M"),
    },
    "periods": periods,
    "competitors": {
        "hkg_today": hkg, "others_avg": others_avg, "bsrm": bsrm,
        "hkg_vs_others": hkg-others_avg, "hkg_vs_bsrm": hkg-bsrm,
    },
    "stocks": stocks,
    "monthly": monthly,
    "top_customers": top_custs,
    "salesmen": salesmen,
}

with open(OUTPUT_FILE, "w") as f:
    json.dump(data, f, separators=(",",":"))

print(f"\n✅ {OUTPUT_FILE} updated successfully!")
print(f"   Latest date: {latest.date()}")
print(f"   MTD production: {periods['mtd']['rod']} MT")
print(f"   MTD sales: {periods['mtd']['sales_vol']} MT")
print(f"   File size: {os.path.getsize(OUTPUT_FILE):,} bytes")
print(f"\nNext step: upload data.json to GitHub")
print(f"  → https://github.com/ahan1712/hkg")
