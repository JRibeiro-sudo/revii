# -*- coding: utf-8 -*-
import warnings, json, sys
warnings.filterwarnings("ignore")
import formulas

PATH="/home/user/revii/financials/Revii_VoiceCare_Financial_Model.xlsx"
xl = formulas.ExcelModel().loads(PATH).finish()
sol = xl.calculate()

vals={}
for k,v in sol.items():
    ku=k.upper()
    if "'!" not in ku: continue
    try:
        val=v.value[0,0]
    except Exception:
        try: val=v.value
        except Exception: continue
    vals[ku]=val

def g(sheet,ref):
    key=f"'[REVII_VOICECARE_FINANCIAL_MODEL.XLSX]{sheet.upper()}'!{ref}"
    return vals.get(key)

# scan for error strings anywhere
errs={}
for k,v in vals.items():
    s=str(v)
    if s.startswith("#") and any(e in s for e in ("REF","NAME","VALUE","DIV","N/A","NUM","NULL")):
        errs.setdefault(s.split("!")[0],[]).append(k)
print("=== ERROR SCAN ===")
if not errs: print("no error values found")
for e,locs in errs.items():
    print(f"{e}: {len(locs)}  e.g. {locs[:6]}")

def row(sheet,label,r,cols="DEFGH",f="{:>14,.0f}"):
    out=[]
    for c in cols:
        v=g(sheet,f"{c}{r}")
        try: out.append(f.format(float(v)))
        except Exception: out.append(f"{str(v)[:14]:>14}")
    print(f"{label:34}"+"".join(out))

print("\n=== ANNUAL (workbook-computed) ===")
print(f"{'':34}"+"".join(f"{'Y'+str(i):>14}" for i in range(1,6)))
for lab,r in [("Net revenue",10),("Cost of service",11),("Gross profit",12),
              ("Total operating expenses",19),("EBITDA",20),("Depreciation",22),
              ("Corporation tax",24),("Net income",25),("Free cash flow",40),
              ("Founders' equity injected",42),("Cum. equity invested",43),
              ("Cash (year end)",44),("Customers PT",29),("Customers ES",30),
              ("Balance check (must be 0)",56)]:
    row("Annual",lab,r)
row("Annual","Gross margin %",13,f="{:>13.1%} ")
row("Annual","EBITDA margin %",21,f="{:>13.1%} ")

print("\n=== DASHBOARD ===")
for lab,ref,f in [("Total equity required","C6","{:,.0f}"),("Deepest cash position","C8","{:,.0f}"),
                  ("EBITDA positive","C12","{}"),("Net income positive","C13","{}"),
                  ("FCF positive","C14","{}"),("Cum FCF above zero","C15","{}"),
                  ("Y5 net revenue","C18","{:,.0f}"),("Y5 gross billings","C19","{:,.0f}"),
                  ("Y5 EBITDA","C20","{:,.0f}"),("Y5 customers total","C24","{:,.0f}"),
                  ("Exit enterprise value","C46","{:,.0f}"),("Exit equity value","C48","{:,.0f}"),
                  ("Money multiple","C49","{:.2f}x"),("IRR","C50","{:.1%}"),("NPV","C51","{:,.0f}")]:
    v=g("Dashboard",ref)
    try: print(f"{lab:34}{f.format(float(v)):>16}")
    except Exception: print(f"{lab:34}{str(v):>16}")

print("\n=== UNIT ECONOMICS (PT Y1 / PT Y5 / ES Y1 / ES Y5) ===")
for lab,r in [("Net revenue per month",29),("Cost of service per month",30),
              ("Gross profit per month",31),("Gross margin",32),("LTV",34),
              ("COCA",35),("LTV / COCA",36),("CAC payback months",37)]:
    row("Dashboard",lab,r,cols="CDEF",f="{:>14,.2f}")

print("\n=== INPUTS spot-check ===")
for lab,ref in [("Net revenue factor PT","C27"),("Net rev/customer Y1 PT","C28"),
                ("Voice minutes/month","C57"),("AI cost/customer Y1","C63"),
                ("Gross margin/customer Y1","C64"),("Monthly churn","C42"),
                ("Avg customer life (mo)","C43"),("PT launch month","C13"),
                ("ES launch month","C14"),("Preventive launch month","C15")]:
    print(f"{lab:34}{str(g('Inputs',ref)):>16}")
