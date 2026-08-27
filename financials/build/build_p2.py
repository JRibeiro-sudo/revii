# -*- coding: utf-8 -*-
"""Part 2: Inputs sheet, with a ref registry the later sheets reuse."""
import json, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

OUT = "/home/user/revii/financials/Revii_VoiceCare_Financial_Model.xlsx"
REG = "/tmp/claude-0/-home-user-revii/e105a84b-100f-5f3c-9085-1a790bbe56dc/scratchpad/refs.json"

FONT="Arial"; BLUE="0000FF"; BLACK="000000"; GREEN="008000"
HDR_FILL=PatternFill("solid",fgColor="1F3864"); SEC_FILL=PatternFill("solid",fgColor="D9E2F3")
YEL_FILL=PatternFill("solid",fgColor="FFFF00")
EUR='#,##0;(#,##0);"-"'; EUR2='#,##0.00;(#,##0.00);"-"'; PCT='0.0%;(0.0%);"-"'
NUM='#,##0;(#,##0);"-"'; NUM1='#,##0.0;(#,##0.0);"-"'; MULT='0.00"x"'; CMIN='0.000"  "'
thin=Side(style="thin",color="D0D0D0"); BOX=Border(bottom=thin)

wb = openpyxl.load_workbook(OUT)
ws = wb.create_sheet("Inputs", 2)
ws.sheet_view.showGridLines = False
for c,w in [("A",3),("B",50),("C",13),("D",13),("E",13),("F",13),("G",13),("H",3),("I",70)]:
    ws.column_dimensions[c].width = w

def st(c,**kw):
    c.font=Font(name=FONT,bold=kw.get("bold",False),color=kw.get("color",BLACK),
                size=kw.get("size",10),italic=kw.get("italic",False))
    if kw.get("fill"): c.fill=kw["fill"]
    if kw.get("fmt"): c.number_format=kw["fmt"]
    if kw.get("align") or kw.get("wrap"):
        c.alignment=Alignment(horizontal=kw.get("align"),vertical="center",wrap_text=kw.get("wrap",False))
    if kw.get("border"): c.border=kw["border"]
    return c
def put(ref,val,**kw):
    ws[ref]=val; return st(ws[ref],**kw)

R = {}
r = [1]
def nxt():
    r[0]+=1; return r[0]

put("B2","VOICE CARE — INPUTS & ASSUMPTIONS",bold=True,size=14,color="1F3864")
put("B3","Blue = edit here.  Green = pulled from the Scenarios tab.  Black = calculated.  Yellow fill = the assumptions worth arguing about.",
    italic=True,size=9,color="595959")
r[0]=4

def section(title):
    n=nxt(); n=nxt()
    ws.cell(n,2,title)
    for c in range(2,8): st(ws.cell(n,c),bold=True,size=10,fill=SEC_FILL,color="1F3864")
    st(ws.cell(n,9),bold=True,size=9,fill=SEC_FILL,color="1F3864")
    ws.cell(n,9,"Source / rationale")
    return n

def scalar(key,label,cval,dval=None,fmt=NUM,note="",color=BLUE,hl=False,dcolor=None):
    n=nxt()
    put(f"B{n}",label,size=10,border=BOX)
    kw=dict(fmt=fmt,align="center",border=BOX,color=color)
    if hl: kw["fill"]=YEL_FILL
    put(f"C{n}",cval,**kw)
    if dval is not None:
        kw2=dict(kw); kw2["color"]=dcolor or color
        put(f"D{n}",dval,**kw2)
    if note: put(f"I{n}",note,size=9,color="595959",wrap=True)
    R[key]=f"$C${n}"; R[key+"_ES"]=f"$D${n}"; R[key+"_row"]=n
    return n

def yeartable(key,label,vals,fmt=PCT,color=BLUE,note="",hl=False):
    n=nxt()
    put(f"B{n}",label,size=10,border=BOX)
    for i,v in enumerate(vals):
        kw=dict(fmt=fmt,align="center",border=BOX,color=color)
        if hl: kw["fill"]=YEL_FILL
        put(f"{chr(67+i)}{n}",v,**kw)
    if note: put(f"I{n}",note,size=9,color="595959",wrap=True)
    R[key]=f"$C${n}:$G${n}"; R[key+"_row"]=n
    return n

def yrhdr(label="Model year"):
    n=nxt()
    put(f"B{n}",label,bold=True,size=9,color="404040")
    for i,t in enumerate(["Year 1","Year 2","Year 3","Year 4","Year 5"]):
        put(f"{chr(67+i)}{n}",t,bold=True,size=9,align="center",color="FFFFFF",fill=HDR_FILL)
    return n

def mkthdr():
    n=nxt()
    put(f"C{n}","Portugal",bold=True,size=9,align="center",color="FFFFFF",fill=HDR_FILL)
    put(f"D{n}","Spain",bold=True,size=9,align="center",color="FFFFFF",fill=HDR_FILL)
    return n

SC = lambda i: f"INDEX(Scenarios!$D${5+i}:$F${5+i},1,$C${R['scen_row']})"

# 1 SCENARIO -------------------------------------------------------------
section("1 · SCENARIO SELECTOR")
n=nxt()
put(f"B{n}","Active scenario  (1 = Conservative, 2 = Base, 3 = Optimistic)",bold=True,size=10,border=BOX)
put(f"C{n}",2,color=BLUE,fmt="0",align="center",fill=YEL_FILL,border=BOX,bold=True)
put(f"I{n}","Change this one cell to reprice the whole business.",size=9,color="595959")
R["scen"]=f"$C${n}"; R["scen_row"]=n
n=nxt()
put(f"B{n}","Scenario in force",size=10,border=BOX)
put(f"C{n}",f"=INDEX(Scenarios!$D$5:$F$5,1,{R['scen']})",color=GREEN,align="center",bold=True,border=BOX)

# 2 TIMING ---------------------------------------------------------------
section("2 · TIMING")
scalar("horizon","Model horizon (months)",60,fmt="0",
       note="Five years. Month 1 is the first month of the venture, not the first month of trading.")
scalar("build","Product build period before launch (months)",4,fmt="0",hl=True,
       note="No revenue during the build. BP tech section sizes the Voice Care build at 3-4 months.")
n=scalar("pt_launch","Portugal launch month",f"=C{R['build_row']}+1",fmt="0",color=BLACK)
scalar("es_launch","Spain launch month",f"=C{R['pt_launch_row']}+12",fmt="0",color=BLACK,
       note="Twelve months after Portugal, per instruction and the deck (Voice Care Spain: +12 months).")
scalar("prev_launch","Preventive Care launch month",f"=C{R['pt_launch_row']}+24",fmt="0",color=BLACK,hl=True,
       note="Deck slide 25: Preventive Care is a post-launch upgrade at +24 months. Pull this forward to test the margin effect.")
scalar("days","Average days per month",30.44,fmt=NUM1,color=BLUE,note="365.25 / 12. Drives voice minutes per month.")

# 3 PRICING --------------------------------------------------------------
section("3 · PRICING, VAT & REVENUE LEAKAGE")
mkthdr()
scalar("price_v","Voice Care price (EUR/month, as paid by customer)",f"={SC(5)}",f"=C{r[0]+1}",
       fmt=EUR2,color=GREEN,dcolor=BLACK,hl=True,
       note="From the Scenarios tab. Spain is priced identically to Portugal by default.")
scalar("price_p","Preventive Care add-on (EUR/month, as paid)",8.99,8.99,fmt=EUR2,
       note="BP Revenue Architecture: Preventive Care add-on at EUR 8.99, available across all tiers.")
scalar("vat_inc","Prices include VAT?  (1 = yes, 0 = no)",1,1,fmt="0",hl=True,
       note="Neither the BP nor the deck mentions VAT. Per instruction the listed price is what the end user pays, so it is treated as VAT-inclusive.")
scalar("vat","VAT rate",0.23,0.21,fmt=PCT,note="Standard rate: Portugal 23%, Spain 21%.")
scalar("price_inf","Annual price increase",0.01,0.01,fmt=PCT,
       note="MBA model escalates Voice Care ~1%/yr (18.99 to 19.76 over five years).")
scalar("leakage","Payment leakage (failed cards, refunds, chargebacks)",0.04,0.04,fmt=PCT,
       note="'Details V2' sheet: 3-5% revenue leakage, recommends modelling a 4% haircut.")
scalar("fee","Billing / payment channel fee",0.025,0.025,fmt=PCT,
       note="Assumes web checkout (Stripe ~2.5%). Selling through the app stores instead would cost 15-30% and break the model — see Dashboard note.")
scalar("net_fact","Net revenue retained per EUR 1 billed",
       f"=IF(C{R['vat_inc_row']}=1,1/(1+C{R['vat_row']}),1)*(1-C{R['leakage_row']})*(1-C{R['fee_row']})",
       f"=IF(D{R['vat_inc_row']}=1,1/(1+D{R['vat_row']}),1)*(1-D{R['leakage_row']})*(1-D{R['fee_row']})",
       fmt=PCT,color=BLACK,note="The gap between headline price and money actually earned.")
scalar("net_arpu","Net revenue per customer, Year 1 (EUR/month)",
       f"=C{R['price_v_row']}*C{R['net_fact_row']}",f"=D{R['price_v_row']}*D{R['net_fact_row']}",
       fmt=EUR2,color=BLACK)

# 4 GROWTH ---------------------------------------------------------------
section("4 · CUSTOMER GROWTH & RETENTION")
mkthdr()
scalar("adds0","New paying customers in the launch month",f"={SC(1)}",f"={SC(2)}",fmt=NUM,color=GREEN,
       note="MBA growth sheet opens Portugal at 10 customers in month 1.")
put(f"B{nxt()}","Monthly growth in new customer additions, by year in that market:",italic=True,size=9,color="595959")
scalar("g1","    Market year 1",0.20,0.20,fmt=PCT,hl=True,
       note="The acquisition ramp. MBA model uses a decaying-rate curve; this is the same shape expressed as a growth rate you can type over.")
scalar("g2","    Market year 2",0.07,0.07,fmt=PCT)
scalar("g3","    Market year 3",0.04,0.04,fmt=PCT)
scalar("g4","    Market year 4",0.025,0.025,fmt=PCT)
scalar("g5","    Market year 5",0.015,0.015,fmt=PCT)
R["growth"]=f"$C${R['g1_row']}:$C${R['g5_row']}"
R["growth_ES"]=f"$D${R['g1_row']}:$D${R['g5_row']}"
scalar("ramp","Growth ramp multiplier (scenario)",f"={SC(3)}",fmt=MULT,color=GREEN,
       note="Scales all five growth rates above at once.")
scalar("cap","Cap on new customers per month",500,1000,fmt=NUM,
       note="A sanity ceiling so an aggressive ramp cannot compound to absurdity.")
scalar("churn","Annual churn rate",f"={SC(4)}",f"=C{r[0]+1}",fmt=PCT,color=GREEN,dcolor=BLACK,hl=True,
       note="MBA model assumes 15%/yr for Voice Care. Base case here is 18%.")
scalar("churn_m","Monthly churn (calculated)",f"=1-(1-C{R['churn_row']})^(1/12)",
       f"=1-(1-D{R['churn_row']})^(1/12)",fmt='0.00%',color=BLACK)
scalar("life","Implied average customer life (months)",f"=1/C{R['churn_m_row']}",
       f"=1/D{R['churn_m_row']}",fmt=NUM1,color=BLACK)
scalar("trial_d","Free trial length (days)",14,14,fmt="0",
       note="BP Layer 1 go-to-market relies on a time-limited free trial. The BP never costs it; this model does.")
scalar("trial_c","Trial to paid conversion",0.35,0.35,fmt=PCT,hl=True,
       note="Drives how many un-monetised trial minutes are burned per paying customer won.")

# 5 PREVENTIVE -----------------------------------------------------------
section("5 · PREVENTIVE CARE ATTACH")
yrhdr("Years since that market launched")
yeartable("attach_pt","Portugal — % of active customers taking Preventive",[0,0,0.10,0.20,0.30],hl=True,
          note="No published attach data exists. These are placeholders and deserve a hard look — Preventive is what repairs the Voice Care margin.")
yeartable("attach_es","Spain — % of active customers taking Preventive",[0,0,0.10,0.20,0.30])
scalar("attach_m","Attach-rate multiplier (scenario)",f"={SC(9)}",fmt=MULT,color=GREEN)
scalar("prev_cost","Preventive Care variable cost (EUR/user/month)",3.60,fmt=EUR2,
       note="BP Financial Plan: EUR 3.60 variable cost, ~60% operating margin on the add-on.")

# 6 VOICE AI -------------------------------------------------------------
section("6 · VOICE AI COST ENGINE  — the core economics")
scalar("calls","AI calls per customer per day",1.0,fmt=NUM1,hl=True,
       note="'Details V2': 1 interaction per day.")
scalar("mins","Minutes per call",f"={SC(6)}",fmt=NUM1,color=GREEN,hl=True,
       note="The MBA sheet says '~5 minutes/day' but its EUR 11.25 cost is 75 min/month = 2.5 min/day. At a true 5 min/day the gross margin goes negative. This is the number to verify in a pilot.")
scalar("mins_mo","Voice minutes per customer per month",
       f"=C{R['calls_row']}*C{R['mins_row']}*C{R['days_row']}",fmt=NUM1,color=BLACK)
scalar("cpm","Cost per voice minute, Year 1 (EUR)",f"={SC(7)}",fmt=CMIN,color=GREEN,hl=True,
       note="MBA model cites Retell AI at EUR 0.15/min.")
scalar("cpm_chg","Annual change in cost per minute",-0.15,fmt=PCT,hl=True,
       note="Inference costs have fallen steadily. A -15%/yr glide is deliberate but assumed, not sourced.")
yrhdr()
yeartable("cpm_yr","Cost per minute by model year (EUR)",
          [f"=$C${R['cpm_row']}"]+[f"={chr(66+i)}{r[0]+1}*(1+$C${R['cpm_chg_row']})" for i in range(1,5)],
          fmt=CMIN,color=BLACK)
scalar("data","Data & storage (EUR/user/month)",0.03,fmt=EUR2,
       note="'Details V2': 72 MB/month of call data at Google Cloud Storage rates.")
scalar("ai_cost1","AI cost per paying customer, Year 1 (EUR/month)",
       f"=C{R['mins_mo_row']}*C{R['cpm_row']}+C{R['data_row']}",fmt=EUR2,color=BLACK,
       note="Compare against net revenue per customer above. That difference is the whole business.")
scalar("gm1","Gross margin per Portugal customer, Year 1",
       f"=IFERROR(1-C{R['ai_cost1_row']}/C{R['net_arpu_row']},0)",fmt=PCT,color=BLACK,hl=True)

# 7 ACQUISITION COST -----------------------------------------------------
section("7 · CUSTOMER ACQUISITION COST")
yrhdr("Years since that market launched")
yeartable("coca_pt","Portugal — cost per acquired customer (EUR)",[38,32,26,22,20],fmt=EUR,
          note="MBA growth sheet: COCA decays from EUR 37.98 to EUR 21.36 as organic and referral channels mature.")
yeartable("coca_es","Spain — cost per acquired customer (EUR)",[38,32,26,22,20],fmt=EUR)
scalar("coca_m","COCA multiplier (scenario)",f"={SC(8)}",fmt=MULT,color=GREEN)

# 8 OPEX -----------------------------------------------------------------
section("8 · OPERATING EXPENSES  (EUR per month, by model year)")
yrhdr()
yeartable("op_found","Founders' salaries (all founders)",[0,0,0,0,0],fmt=EUR,hl=True,
          note="Zero, per instruction. A normally-paid founding team would add roughly EUR 150-250k a year from Year 2 and would materially change the equity requirement.")
yeartable("op_dev","Developer / platform maintenance",[1500,4000,8000,12000,16000],fmt=EUR,
          note="'Details V2' lean case: EUR 1,500-4,500/month part-time maintenance. Scaled up as the base grows.")
yeartable("op_sup","Customer support",[0,1200,3000,5000,7000],fmt=EUR,
          note="'Details V2': EUR 800-1,200/month part-time to start. BP has support part-time from Year 2.")
yeartable("op_mkt","Brand marketing & content (excludes COCA)",[1000,2500,5000,7000,9000],fmt=EUR,hl=True,
          note="Deliberately a fraction of the BP's EUR 500k Year 1 brand budget, which assumed a funded launch. Multiplied by the scenario marketing lever.")
yeartable("op_cloud","Cloud & tools (non-usage infrastructure)",[250,400,700,1000,1400],fmt=EUR,
          note="'Details V2': GCP base, monitoring and logging. Usage-driven AI cost sits in COGS, not here.")
yeartable("op_ga","Accounting, admin & legal",[400,500,700,900,1100],fmt=EUR)
yeartable("op_ins","Insurance (professional liability + cyber)",[0,625,625,750,900],fmt=EUR,
          note="'Details V2': EUR 4,000 liability + EUR 3,500 cyber per year, from the first year of trading.")
yeartable("op_off","Office / co-working",[0,300,600,900,1200],fmt=EUR,
          note="Remote to start; BP assumes EUR 3,000/month office, which a founder-led launch does not need.")
yeartable("op_es","Spain local lead & market management",[0,1500,3000,4000,5000],fmt=EUR,
          note="BP Place section: marketing must be run locally in each market. Charged only from the Spain launch month.")
yeartable("op_other","Other / contingency",[200,400,800,1200,1600],fmt=EUR)

# 9 CAPEX ----------------------------------------------------------------
section("9 · CAPITAL EXPENDITURE & DEPRECIATION")
scalar("capex_v","Voice Care platform build (EUR)",46000,fmt=EUR,hl=True,
       note="MBA model, Voice Care R&D line. An AI-native build was costed at EUR 38,000 in 'Details V2'. Spread evenly across the build months.")
scalar("capex_p","Preventive Care module build (EUR)",24250,fmt=EUR,
       note="MBA model, Preventive Care CAPEX. Incurred in its launch month.")
scalar("capex_es","Spain localisation & compliance (EUR)",10000,fmt=EUR,
       note="Not in the MBA model. Added here: Spanish-language voice models, legal, local payment setup.")
scalar("capex_ong","Ongoing development capex, % of cumulative build per year",0.20,fmt=PCT,
       note="MBA model applies 20% of initial R&D per year as maintenance capex.")
scalar("dep_yrs","Depreciation period (years, straight line)",5,fmt="0",
       note="MBA model depreciates R&D over 5 years.")

# 10 WORKING CAPITAL -----------------------------------------------------
section("10 · WORKING CAPITAL & CASH")
scalar("recv_d","Receivable days",5,fmt="0",note="Card-on-file subscriptions settle in days, not months.")
scalar("pay_d","Payable days",30,fmt="0",note="Applied to cost of service and cash operating costs.")
scalar("prepay","% of new customers taking an annual prepaid plan",0.0,fmt=PCT,hl=True,
       note="A real self-funding lever: annual prepay pulls twelve months of cash forward and cuts the equity needed. Base case is 0% so the result is not flattered.")
scalar("prepay_disc","Discount offered on the annual prepaid plan",0.15,fmt=PCT)
scalar("mincash","Minimum cash buffer to hold (EUR)",15000,fmt=EUR,
       note="Founders top up to this level each month. Sets the floor for the equity requirement.")

# 11 TAX / RETURN --------------------------------------------------------
section("11 · TAX, FUNDING & RETURN")
scalar("tax","Corporate tax rate",0.19,fmt=PCT,note="MBA model assumes 19%. Losses are carried forward without limit.")
scalar("disc","Discount rate for NPV",0.12,fmt=PCT,hl=True,
       note="MBA sheets use 12% (older) and 11.5% WACC (DCF). For an unfunded pre-revenue venture this is generous.")
scalar("exit_basis","Exit basis  (1 = multiple of Year 5 revenue, 2 = multiple of Year 5 EBITDA)",1,fmt="0")
scalar("exit_rev","Exit multiple — times Year 5 net revenue",3.0,fmt=MULT,hl=True,
       note="A convention for a subscription business, not a valuation. Change it and the IRR changes with it.")
scalar("exit_ebitda","Exit multiple — times Year 5 EBITDA",8.0,fmt=MULT)

put(f"B{nxt()+1}","Funding structure: 100% founders' equity — no debt, no grant, no external round.",
    bold=True,size=10,color="1F3864")
put(f"B{nxt()}","Equity is injected month by month, only as much as is needed to hold the minimum cash buffer. The total injected is the investment figure on the Dashboard.",
    italic=True,size=9,color="595959")

wb.save(OUT)
json.dump(R, open(REG,"w"), indent=1)
print("part 2 saved. rows used:", r[0], "| refs:", len(R))
