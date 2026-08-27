# -*- coding: utf-8 -*-
"""Part 5: break-even helper rows, Annual statements, Dashboard."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUT="/home/user/revii/financials/Revii_VoiceCare_Financial_Model.xlsx"
FONT="Arial"; BLACK="000000"; GREEN="008000"; BLUE="0000FF"
HDR=PatternFill("solid",fgColor="1F3864"); SEC=PatternFill("solid",fgColor="D9E2F3")
TOT=PatternFill("solid",fgColor="E2EFDA"); WARN=PatternFill("solid",fgColor="FCE4D6")
TILE=PatternFill("solid",fgColor="EDF3FA")
EUR='#,##0;(#,##0);"-"'; EUR2='#,##0.00;(#,##0.00);"-"'; PCT='0.0%;(0.0%);"-"'
NUM='#,##0;(#,##0);"-"'; NUM1='#,##0.0;(#,##0.0);"-"'; MULT='0.00"x"'
thin=Side(style="thin",color="BFBFBF"); BOX=Border(top=thin,bottom=thin,left=thin,right=thin)

N=60
def C(i): return get_column_letter(3+i)
YR=[("D","O"),("P","AA"),("AB","AM"),("AN","AY"),("AZ","BK")]

wb=openpyxl.load_workbook(OUT)

def st(c,**kw):
    c.font=Font(name=FONT,bold=kw.get("bold",False),color=kw.get("color",BLACK),
                size=kw.get("size",10),italic=kw.get("italic",False))
    if kw.get("fill"): c.fill=kw["fill"]
    if kw.get("fmt"): c.number_format=kw["fmt"]
    if kw.get("align") or kw.get("wrap"):
        c.alignment=Alignment(horizontal=kw.get("align"),vertical="center",wrap_text=kw.get("wrap",False))
    if kw.get("border"): c.border=kw["border"]
    return c

# ---------------------------------------------- helper rows on P&L Monthly
m=wb["P&L Monthly"]
def mput(ref,v,**kw):
    m[ref]=v; return st(m[ref],**kw)
mput("B78","BREAK-EVEN HELPERS  (first month from which the measure stays positive)",bold=True,size=9,fill=SEC,color="1F3864")
for c in range(3,3+N): st(m.cell(78,c),fill=SEC)
helpers=[(79,"EBITDA turns and stays positive",23),(80,"Free cash flow turns and stays positive",59),
         (81,"Net income turns and stays positive",29),(82,"Cumulative free cash flow turns positive",60)]
for row,label,src in helpers:
    mput(f"B{row}",label,size=9)
    mput(f"C{row}","mo",size=8,color="808080",align="center")
    for i in range(1,N+1):
        col=C(i)
        f=(f"=IF(AND({col}{src}>0,MIN({col}{src}:$BK${src})>0),{col}4,9999)" if row!=82
           else f"=IF({col}60>0,{col}4,9999)")
        mput(f"{col}{row}",f,fmt="0",size=9)

# ---------------------------------------------------------------- ANNUAL
a=wb.create_sheet("Annual",6)
a.sheet_view.showGridLines=False
a.column_dimensions["A"].width=2
a.column_dimensions["B"].width=44
a.column_dimensions["C"].width=7
for col in "DEFGH": a.column_dimensions[col].width=15
a.column_dimensions["I"].width=16
def aput(ref,v,**kw):
    a[ref]=v; return st(a[ref],**kw)
def aband(row,label):
    a.cell(row,2,label)
    for c in range(2,10): st(a.cell(row,c),bold=True,size=10,fill=SEC,color="1F3864")

aput("B2","ANNUAL FINANCIAL STATEMENTS  —  YEARS 1 TO 5",bold=True,size=14,color="1F3864")
aput("B3","All figures in euros. Year 1 includes the product build months, so it carries only part of a year of trading.",
     italic=True,size=9,color="595959")
aput("B5","",fill=HDR); aput("C5","Unit",bold=True,size=9,color="FFFFFF",fill=HDR,align="center")
for i,t in enumerate(["Year 1","Year 2","Year 3","Year 4","Year 5"]):
    aput(f"{chr(68+i)}5",t,bold=True,size=10,color="FFFFFF",fill=HDR,align="center")
aput("I5","5-year total",bold=True,size=10,color="FFFFFF",fill=HDR,align="center")

def arow(row,label,src,kind="flow",fmt=EUR,*,bold=False,fill=None,ind=0,total=True,sheet="P&L Monthly"):
    aput(f"B{row}","    "*ind+label,size=10,bold=bold,fill=fill)
    aput(f"C{row}","EUR" if fmt in (EUR,EUR2) else ("%" if fmt==PCT else ""),
         size=8,color="808080",align="center",fill=fill)
    for i,(s,e) in enumerate(YR):
        col=chr(68+i)
        f=f"=SUM('{sheet}'!{s}{src}:{e}{src})" if kind=="flow" else f"='{sheet}'!{e}{src}"
        aput(f"{col}{row}",f,fmt=fmt,size=10,bold=bold,fill=fill,color=GREEN)
    if total and kind=="flow":
        aput(f"I{row}",f"=SUM(D{row}:H{row})",fmt=fmt,size=10,bold=True,fill=fill)
    elif total:
        aput(f"I{row}",f"=H{row}",fmt=fmt,size=10,bold=True,fill=fill,color="808080")

def acalc(row,label,fmla,fmt=PCT,*,bold=False,fill=None,ind=0,tot=None):
    aput(f"B{row}","    "*ind+label,size=10,bold=bold,fill=fill)
    aput(f"C{row}","%" if fmt==PCT else "",size=8,color="808080",align="center",fill=fill)
    for i in range(5):
        col=chr(68+i)
        aput(f"{col}{row}",fmla.replace("{c}",col),fmt=fmt,size=10,bold=bold,fill=fill)
    if tot: aput(f"I{row}",tot,fmt=fmt,size=10,bold=True,fill=fill)

aband(7,"PROFIT & LOSS")
arow(8,"Net revenue — Portugal",8,ind=1)
arow(9,"Net revenue — Spain",9,ind=1)
arow(10,"Total net revenue",10,bold=True,fill=TOT)
arow(11,"Cost of service",13)
arow(12,"Gross profit",14,bold=True)
acalc(13,"Gross margin","=IFERROR({c}12/{c}10,0)",tot="=IFERROR(I12/I10,0)")
arow(14,"Customer acquisition (COCA)",17,ind=1)
arow(15,"Brand marketing & content",18,ind=1)
arow(16,"Personnel",19,ind=1)
arow(17,"Cloud & tools",20,ind=1)
arow(18,"General & administrative",21,ind=1)
arow(19,"Total operating expenses",22,bold=True)
arow(20,"EBITDA",23,bold=True,fill=TOT)
acalc(21,"EBITDA margin","=IFERROR({c}20/{c}10,0)",tot="=IFERROR(I20/I10,0)")
arow(22,"Depreciation",25)
arow(23,"EBIT",26,bold=True)
arow(24,"Corporation tax",28)
arow(25,"Net income",29,bold=True,fill=TOT)
acalc(26,"Net margin","=IFERROR({c}25/{c}10,0)",tot="=IFERROR(I25/I10,0)")

aband(28,"CUSTOMERS")
arow(29,"Active customers — Portugal (year end)",13,"stock",NUM,ind=1,sheet="Model PT")
arow(30,"Active customers — Spain (year end)",13,"stock",NUM,ind=1,sheet="Model ES")
acalc(31,"Total active customers (year end)","={c}29+{c}30",NUM,bold=True,fill=TOT,tot="=H31")
aput("B32","    New customers acquired during the year",size=10)
aput("C32","",size=8)
for i,(s,e) in enumerate(YR):
    aput(f"{chr(68+i)}32",f"=SUM('Model PT'!{s}11:{e}11)+SUM('Model ES'!{s}11:{e}11)",fmt=NUM,size=10,color=GREEN)
aput("I32","=SUM(D32:H32)",fmt=NUM,size=10,bold=True)
aput("B33","    Net revenue per customer per month (blended)",size=10)
for i,(s,e) in enumerate(YR):
    aput(f"{chr(68+i)}33",
         f"=IFERROR({chr(68+i)}10/(SUM('Model PT'!{s}14:{e}14)+SUM('Model ES'!{s}14:{e}14)),0)",
         fmt=EUR2,size=10)

aband(35,"CASH FLOW")
arow(36,"EBITDA",55,ind=1)
arow(37,"Corporation tax paid",56,ind=1)
arow(38,"Capital expenditure",57,ind=1)
arow(39,"Change in working capital",58,ind=1)
arow(40,"Free cash flow",59,bold=True,fill=TOT)
arow(41,"Cumulative free cash flow (year end)",60,"stock",bold=True)
arow(42,"Founders' equity injected",62,bold=True,fill=WARN)
arow(43,"Cumulative equity invested (year end)",63,"stock",bold=True,fill=WARN)
arow(44,"Cash — closing balance",64,"stock",bold=True,fill=TOT)

aband(46,"BALANCE SHEET  (year end)")
arow(47,"Cash",67,"stock",ind=1)
arow(48,"Receivables",68,"stock",ind=1)
arow(49,"Net fixed assets",69,"stock",ind=1)
arow(50,"Total assets",70,"stock",bold=True)
arow(51,"Payables",71,"stock",ind=1)
arow(52,"Deferred revenue",72,"stock",ind=1)
arow(53,"Paid-in equity",73,"stock",ind=1)
arow(54,"Retained earnings",74,"stock",ind=1)
arow(55,"Total liabilities & equity",75,"stock",bold=True)
arow(56,"Balance check (must be nil)",76,"stock",bold=True)

# ---------------------------------------------------------------- DASHBOARD
d=wb.create_sheet("Dashboard",1)
d.sheet_view.showGridLines=False
d.column_dimensions["A"].width=2
d.column_dimensions["B"].width=46
d.column_dimensions["C"].width=18
d.column_dimensions["D"].width=14
for col in "EFGH": d.column_dimensions[col].width=13
d.column_dimensions["I"].width=3
d.column_dimensions["J"].width=64
def dput(ref,v,**kw):
    d[ref]=v; return st(d[ref],**kw)
def dband(row,label,span=8):
    d.cell(row,2,label)
    for c in range(2,2+span): st(d.cell(row,c),bold=True,size=10,fill=SEC,color="1F3864")

dput("B2","VOICE CARE  —  DASHBOARD",bold=True,size=16,color="1F3864")
dput("B3",'=\"Scenario in force:  \"&Inputs!C8&"   ·   change it on Inputs!C7 (1 Conservative / 2 Base / 3 Optimistic)"',
     bold=True,size=10,color=GREEN)

def kpi(row,label,fmla,fmt,note="",big=False):
    dput(f"B{row}",label,size=10,bold=big,border=BOX)
    dput(f"C{row}",fmla,fmt=fmt,size=12 if big else 10,bold=True,align="center",
         border=BOX,fill=TILE if big else None)
    if note: dput(f"J{row}",note,size=9,color="595959",wrap=True)

dband(5,"THE INVESTMENT")
kpi(6,"Total founders' equity required","=MAX('P&L Monthly'!$D$63:$BK$63)",EUR,
    "Peak cumulative cash the founders must put in. This is the number to raise among yourselves.",big=True)
kpi(7,"Of which, spent before Portugal launches","='P&L Monthly'!"+YR[0][0]+"63",EUR,
    "Equity injected in month 1 — the build period. See Inputs section 2 for the build length.")
kpi(8,"Deepest cash position (cumulative free cash flow)","=MIN('P&L Monthly'!$D$60:$BK$60)",EUR,
    "The trough. The gap between this and zero is what the business has to earn back.")
kpi(9,"Month of the deepest cash position",
    '=\"Month \"&INDEX(\'P&L Monthly\'!$D$4:$BK$4,1,MATCH(MIN(\'P&L Monthly\'!$D$60:$BK$60),\'P&L Monthly\'!$D$60:$BK$60,0))',"General")

dband(11,"WHEN IT TURNS")
for row,label,hrow,note in [
    (12,"EBITDA turns positive (and stays)",79,"Operating profit before depreciation."),
    (13,"Net income turns positive (and stays)",81,""),
    (14,"Free cash flow turns positive (and stays)",80,"The month the business stops consuming founders' money."),
    (15,"Cumulative free cash flow back above zero",82,"The month the founders are, in cash terms, whole again.")]:
    kpi(row,label,
        f'=IF(MIN(\'P&L Monthly\'!$D${hrow}:$BK${hrow})>=9999,"beyond year 5","Month "&MIN(\'P&L Monthly\'!$D${hrow}:$BK${hrow}))',
        "General",note)

dband(17,"SCALE AT YEAR 5")
kpi(18,"Net revenue","=Annual!H10",EUR,"After VAT, payment leakage and channel fees.")
kpi(19,"Gross billings (what customers pay)","=SUM('Model PT'!AZ23:BK23)+SUM('Model ES'!AZ23:BK23)",EUR,
    "The headline number. The gap to net revenue is VAT and leakage.")
kpi(20,"EBITDA","=Annual!H20",EUR)
kpi(21,"EBITDA margin","=Annual!H21",PCT)
kpi(22,"Active customers — Portugal","=Annual!H29",NUM)
kpi(23,"Active customers — Spain","=Annual!H30",NUM)
kpi(24,"Active customers — total","=Annual!H31",NUM,big=True)
kpi(25,"Cumulative net income, years 1-5","=Annual!I25",EUR)

dband(27,"UNIT ECONOMICS  (per customer)")
dput("C27","",fill=SEC)
for i,t in enumerate(["Portugal Y1","Portugal Y5","Spain Y1","Spain Y5"]):
    dput(f"{chr(67+i)}28",t,bold=True,size=9,color="FFFFFF",fill=HDR,align="center")
UE=[("Net revenue per month",
     ["=IFERROR(SUM('Model PT'!D24:O25)/SUM('Model PT'!D14:O14),0)",
      "=IFERROR(SUM('Model PT'!AZ24:BK25)/SUM('Model PT'!AZ14:BK14),0)",
      "=IFERROR(SUM('Model ES'!D24:O25)/SUM('Model ES'!D14:O14),0)",
      "=IFERROR(SUM('Model ES'!AZ24:BK25)/SUM('Model ES'!AZ14:BK14),0)"],EUR2),
    ("Cost of service per month",
     ["=IFERROR(SUM('Model PT'!D34:O34)/SUM('Model PT'!D14:O14),0)",
      "=IFERROR(SUM('Model PT'!AZ34:BK34)/SUM('Model PT'!AZ14:BK14),0)",
      "=IFERROR(SUM('Model ES'!D34:O34)/SUM('Model ES'!D14:O14),0)",
      "=IFERROR(SUM('Model ES'!AZ34:BK34)/SUM('Model ES'!AZ14:BK14),0)"],EUR2),
    ("Gross profit per month",["=C29-C30","=D29-D30","=E29-E30","=F29-F30"],EUR2),
    ("Gross margin",["=IFERROR(C31/C29,0)","=IFERROR(D31/D29,0)","=IFERROR(E31/E29,0)","=IFERROR(F31/F29,0)"],PCT),
    ("Average customer life (months)",["=Inputs!$C$43","=Inputs!$C$43","=Inputs!$D$43","=Inputs!$D$43"],NUM1),
    ("Lifetime value (gross profit basis)",["=C31*C33","=D31*D33","=E31*E33","=F31*F33"],EUR),
    ("Cost per acquired customer",
     ["=Inputs!$C$68*Inputs!$C$70","=Inputs!$G$68*Inputs!$C$70",
      "=Inputs!$C$69*Inputs!$C$70","=Inputs!$G$69*Inputs!$C$70"],EUR2),
    ("LTV / COCA",["=IFERROR(C34/C35,0)","=IFERROR(D34/D35,0)","=IFERROR(E34/E35,0)","=IFERROR(F34/F35,0)"],MULT),
    ("Months to pay back the acquisition cost",
     ["=IFERROR(C35/C31,0)","=IFERROR(D35/D31,0)","=IFERROR(E35/E31,0)","=IFERROR(F35/F31,0)"],NUM1)]
r=29
for label,fs,fmt in UE:
    dput(f"B{r}",label,size=10,border=BOX)
    for i,f in enumerate(fs):
        dput(f"{chr(67+i)}{r}",f,fmt=fmt,size=10,align="center",border=BOX,
             bold=(label=="LTV / COCA"))
    r+=1
dput("J29","Net revenue is after VAT, leakage and channel fees, so these are the real economics — not the headline price.",
     size=9,color="595959",wrap=True)
dput("J34","A LTV/COCA above 3.0x is the usual threshold for a healthy subscription business.",
     size=9,color="595959",wrap=True)

dband(39,"RETURN ON THE FOUNDERS' MONEY")
dput("B40","Annual equity cash flow (negative = money in)",bold=True,size=10)
for i,t in enumerate(["Year 1","Year 2","Year 3","Year 4","Year 5","Exit"]):
    dput(f"{chr(67+i)}40",t,bold=True,size=9,color="FFFFFF",fill=HDR,align="center")
for i in range(5):
    dput(f"{chr(67+i)}41",f"=-Annual!{chr(68+i)}42",fmt=EUR,size=10,align="center",border=BOX)
dput("H41","=C48",fmt=EUR,size=10,align="center",border=BOX,bold=True)
dput("B41","Equity in / exit proceeds out",size=10,border=BOX)

kpi(43,"Total founders' equity invested","=SUM(Annual!D42:H42)",EUR)
kpi(44,"Year 5 net revenue","=Annual!H10",EUR)
kpi(45,"Year 5 EBITDA","=Annual!H20",EUR)
kpi(46,"Exit enterprise value",
    "=IF(Inputs!$C$102=1,Annual!H10*Inputs!$C$103,Annual!H20*Inputs!$C$104)",EUR,
    "A convention for framing the return, not a valuation. Set the basis and multiple on Inputs section 11.")
kpi(47,"Plus year 5 cash, less debt","=Annual!H47",EUR,"No debt in this plan — funding is 100% founders' equity.")
kpi(48,"Exit equity value","=MAX(0,C46+C47)",EUR,big=True)
kpi(49,"Money multiple on invested equity","=IFERROR(C48/C43,0)",MULT,big=True)
kpi(50,"IRR to the founders","=IFERROR(IRR(C41:H41),0)",PCT,
    "Injections are treated as start-of-year, the exit as end of year 5.",big=True)
kpi(51,"NPV of the equity cash flows","=C41+NPV(Inputs!$C$101,D41:H41)",EUR,
    "Discounted at the rate on Inputs!C101.")

dband(53,"YEAR BY YEAR")
dput("B54","",fill=HDR)
for i,t in enumerate(["Year 1","Year 2","Year 3","Year 4","Year 5"]):
    dput(f"{chr(67+i)}54",t,bold=True,size=9,color="FFFFFF",fill=HDR,align="center")
mini=[("Net revenue",10),("EBITDA",20),("Net income",25),("Free cash flow",40),
      ("Founders' equity injected",42),("Cash at year end",47),("Active customers",31)]
r=55
for label,src in mini:
    dput(f"B{r}",label,size=10,border=BOX)
    for i in range(5):
        dput(f"{chr(67+i)}{r}",f"=Annual!{chr(68+i)}{src}",
             fmt=NUM if label=="Active customers" else EUR,size=10,align="center",border=BOX)
    r+=1

dband(63,"READ THIS BEFORE QUOTING ANY NUMBER ABOVE",8)
notes=[
 ("VAT is the biggest change from the MBA model",
  "Neither the plan nor the deck mentions VAT. Treating the listed price as VAT-inclusive removes about 19% of headline revenue in Portugal before any cost. If Revii intends to quote prices ex-VAT instead, set Inputs!C22 and D22 to 0 and every figure above improves sharply."),
 ("Voice minutes are the whole margin",
  "The MBA cost of EUR 11.28 per user implies 2.5 minutes a day, not the 5 minutes its own note describes. Doubling minutes roughly doubles cost of service and can take gross margin negative. Measure real call length in a pilot before committing to the price."),
 ("Founders are unpaid in every year",
  "Per instruction. Paying a founding team properly would add roughly EUR 150-250k a year from year 2 and would push break-even out materially."),
 ("Selling through the app stores would break this",
  "The model assumes web checkout at a 2.5% fee. Apple and Google take 15-30%, which is larger than the entire gross margin at the base price. Bill on the web."),
 ("Preventive Care arrives late",
  "The deck puts it 24 months after launch. It carries a far better margin than Voice Care, so pulling it forward on Inputs!C15 is one of the cheapest ways to improve this plan."),
 ("Working capital is approximated",
  "Receivables and payables use days ratios; VAT settlement timing is ignored. Fine for sizing, not for a cash-management calendar."),
]
r=64
for h,b in notes:
    dput(f"B{r}",h,bold=True,size=10)
    dput(f"C{r}",b,size=9,color="595959",wrap=True)
    d.merge_cells(f"C{r}:H{r}")
    d.row_dimensions[r].height=32
    r+=1

wb.save(OUT)
print("part 5 saved")
