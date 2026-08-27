# -*- coding: utf-8 -*-
"""Revii Voice Care standalone model - part 1: Cover, Scenarios, Inputs."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUT = "/home/user/revii/financials/Revii_VoiceCare_Financial_Model.xlsx"

FONT = "Arial"
BLUE = "0000FF"      # hardcoded input
BLACK = "000000"     # formula
GREEN = "008000"     # link to another sheet
YELLOW = "FFFF00"    # key assumption to fill
HDR_FILL = PatternFill("solid", fgColor="1F3864")
SEC_FILL = PatternFill("solid", fgColor="D9E2F3")
SUB_FILL = PatternFill("solid", fgColor="F2F2F2")
TOT_FILL = PatternFill("solid", fgColor="E2EFDA")
YEL_FILL = PatternFill("solid", fgColor=YELLOW)

EUR   = '#,##0;(#,##0);"-"'
EUR2  = '#,##0.00;(#,##0.00);"-"'
PCT   = '0.0%;(0.0%);"-"'
PCT2  = '0.00%;(0.00%);"-"'
NUM   = '#,##0;(#,##0);"-"'
NUM1  = '#,##0.0;(#,##0.0);"-"'
MULT  = '0.00"x"'

thin = Side(style="thin", color="BFBFBF")
BOX = Border(top=thin, bottom=thin, left=thin, right=thin)
TOPLINE = Border(top=Side(style="thin", color="808080"))

wb = openpyxl.Workbook()

def style(c, *, bold=False, color=BLACK, size=10, fill=None, fmt=None,
          align=None, italic=False, border=None, wrap=False):
    c.font = Font(name=FONT, bold=bold, color=color, size=size, italic=italic)
    if fill: c.fill = fill
    if fmt: c.number_format = fmt
    if align or wrap: c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if border: c.border = border
    return c

def put(ws, ref, value, **kw):
    ws[ref] = value
    return style(ws[ref], **kw)

def section(ws, row, text, span=8):
    ws.cell(row, 1, text)
    for c in range(1, span + 1):
        style(ws.cell(row, c), bold=True, size=10, fill=SEC_FILL, color="1F3864")
    return row

# ---------------------------------------------------------------- COVER
ws = wb.active
ws.title = "README"
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 3
ws.column_dimensions["B"].width = 46
ws.column_dimensions["C"].width = 92

put(ws, "B2", "REVII  ·  VOICE CARE STANDALONE FINANCIAL MODEL", bold=True, size=16, color="1F3864")
put(ws, "B3", "Five-year plan for a real launch of Voice Care only — Portugal first, Spain from month 17", italic=True, size=10, color="595959")

rows = [
    ("", ""),
    ("PURPOSE", ""),
    ("What this is", "A practical sizing model for launching Voice Care as a standalone business: how much founders' equity it needs, when it turns cash-positive, and what the return looks like. It is a decision tool, not a statutory set of accounts."),
    ("What it is not", "Not an accounting model. No deferred tax, no lease accounting, no monthly VAT settlement timing. Working capital is approximated with days ratios."),
    ("", ""),
    ("HOW TO USE IT", ""),
    ("1. Pick a scenario", "On 'Inputs' cell C6, enter 1 (Conservative), 2 (Base) or 3 (Optimistic). Every sheet updates. The three input sets live on the 'Scenarios' tab and can be edited there."),
    ("2. Override anything", "Every blue cell on 'Inputs' is editable. Blue cells fed by the scenario selector are on 'Scenarios'; all other blue cells are direct inputs."),
    ("3. Read the answer", "'Dashboard' carries the headline: equity required, break-even month, Year 5 scale, and the return on the founders' money."),
    ("", ""),
    ("COLOUR LEGEND", ""),
    ("Blue text", "A hardcoded input. Change these."),
    ("Black text", "A formula. Do not overwrite."),
    ("Green text", "A link pulled from another sheet."),
    ("Yellow fill", "A key assumption that most deserves challenge."),
    ("", ""),
    ("STRUCTURE", ""),
    ("Scenarios", "The three scenario input sets. Ten levers that swing the business."),
    ("Inputs", "Every other assumption, with the source of each one noted alongside."),
    ("Model PT / Model ES", "The monthly engine for each market: customers, revenue, cost of service, acquisition. Kept separate so the two revenue streams are never blended."),
    ("P&L Monthly", "60 months consolidated: P&L, working capital, capex, cash flow, and the founders' equity injection."),
    ("Annual", "Years 1-5 P&L, cash flow and a light balance sheet."),
    ("Dashboard", "Headline KPIs, unit economics and the return calculation."),
    ("", ""),
    ("KEY DECISIONS BEHIND THE BASE CASE", ""),
    ("Prices are VAT-inclusive", "Neither the business plan nor the deck mentions VAT anywhere. Per instruction, the listed price is what the customer pays, so net revenue is the price divided by (1 + VAT). This costs ~19% of headline revenue in Portugal and is the single largest correction against the MBA model."),
    ("Voice minutes drive everything", "The MBA model's EUR 11.28 cost per user is 75 minutes/month at EUR 0.15/min, i.e. 2.5 min/day — even though its own note says '~5 minutes/day'. At 5 min/day the product has a negative gross margin. Base case uses 1 call/day x 2.5 min. Test this first."),
    ("Founders take no salary", "Per instruction. OPEX therefore excludes founder compensation in all five years, which flatters EBITDA relative to a normally-staffed company."),
    ("100% founders' equity", "No debt, no grant. The model injects equity monthly, only as needed to hold the minimum cash buffer. Total injected is the investment figure."),
    ("Portugal builds first", "Months 1-4 are product build with no revenue. Portugal launches month 5, Spain month 17 (twelve months later), Preventive Care month 29 (+24 months, per the deck)."),
    ("", ""),
    ("SOURCES", ""),
    ("Financials_Revii_v3_full_product.xlsx", "Voice Care price, COGS build-up, COCA, churn, CAPEX, growth curve shape, OPEX structure."),
    ("Lisbon MBA BP report (.docx)", "Preventive Care pricing and cost, market sequencing, operating model, team structure."),
    ("Final version deck (.pdf)", "Launch sequencing (Spain +12 months, Preventive +24 months), unit economics per tier. Where the deck and the Excel disagree, the deck wins."),
]
r = 5
for label, body in rows:
    if label and not body:
        put(ws, f"B{r}", label, bold=True, size=11, color="1F3864")
        ws.merge_cells(f"B{r}:C{r}")
    elif label:
        put(ws, f"B{r}", label, bold=True, size=10)
        put(ws, f"C{r}", body, size=10, wrap=True)
        ws.row_dimensions[r].height = 28 if len(body) > 110 else 14
    r += 1

put(ws, "B" + str(r + 1), "Built for the Revii founding team · all figures in euros unless stated", italic=True, size=9, color="808080")

# ---------------------------------------------------------------- SCENARIOS
sc = wb.create_sheet("Scenarios")
sc.sheet_view.showGridLines = False
sc.column_dimensions["A"].width = 3
sc.column_dimensions["B"].width = 44
sc.column_dimensions["C"].width = 10
for c in "DEF": sc.column_dimensions[c].width = 15
sc.column_dimensions["G"].width = 3
sc.column_dimensions["H"].width = 62

put(sc, "B2", "SCENARIO LEVERS", bold=True, size=14, color="1F3864")
put(sc, "B3", "The ten assumptions that actually move this business. Edit any cell below; the model uses the column chosen on Inputs!C6.", italic=True, size=9, color="595959")

put(sc, "B5", "Lever", bold=True, fill=HDR_FILL, color="FFFFFF")
put(sc, "C5", "Unit", bold=True, fill=HDR_FILL, color="FFFFFF", align="center")
put(sc, "D5", "1 Conservative", bold=True, fill=HDR_FILL, color="FFFFFF", align="center")
put(sc, "E5", "2 Base", bold=True, fill=HDR_FILL, color="FFFFFF", align="center")
put(sc, "F5", "3 Optimistic", bold=True, fill=HDR_FILL, color="FFFFFF", align="center")
put(sc, "H5", "Why it matters / where the base case comes from", bold=True, fill=HDR_FILL, color="FFFFFF")

SCEN = [
    ("New paying customers in Portugal launch month", "cust", 6, 10, 16, NUM,
     "MBA model starts Portugal at 10 in month 1. Conservative assumes a slower cold start."),
    ("New paying customers in Spain launch month", "cust", 8, 15, 25, NUM,
     "Spain starts warm: brand, content and playbook already exist from Portugal."),
    ("Growth ramp multiplier on monthly new-customer growth", "x", 0.70, 1.00, 1.35, MULT,
     "Scales the whole acquisition ramp set on Inputs rows 30-34. 1.00x reproduces the base ramp."),
    ("Annual churn rate", "% / yr", 0.25, 0.18, 0.12, PCT,
     "MBA model assumes 15%/yr. For a EUR 19 consumer subscription that is optimistic, so base is 18%."),
    ("Voice Care price (EUR/month, incl. VAT)", "EUR", 18.99, 18.99, 24.99, EUR2,
     "Deck price is 18.99. Survey median willingness to pay for a basic tier was EUR 26 (mean 31), so the optimistic case tests 24.99."),
    ("Minutes per AI call", "min", 3.0, 2.5, 2.0, NUM1,
     "The MBA COGS of EUR 11.28 implies 2.5 min/day at 1 call/day, despite its note saying '~5 min/day'. This is the assumption to challenge hardest."),
    ("Cost per AI voice minute, Year 1", "EUR", 0.18, 0.15, 0.08, '0.000"  "',
     "MBA model cites Retell at EUR 0.15/min. Optimistic reflects volume pricing and cheaper models."),
    ("COCA multiplier on the acquisition cost curve", "x", 1.30, 1.00, 0.80,  MULT,
     "Scales the cost-per-acquired-customer curve on Inputs rows 62-63."),
    ("Preventive Care attach-rate multiplier", "x", 0.60, 1.00, 1.40, MULT,
     "Scales the attach curve on Inputs rows 44-45. Preventive is the margin repair, so attach matters more than it looks."),
    ("Brand marketing spend multiplier", "x", 1.00, 1.00, 1.20, MULT,
     "Scales the brand/content line in OPEX. Higher spend in the optimistic case funds the faster ramp."),
]
r = 6
for name, unit, lo, mid, hi, fmt, why in SCEN:
    put(sc, f"B{r}", name, size=10, border=BOX)
    put(sc, f"C{r}", unit, size=9, color="808080", align="center", border=BOX)
    for col, val in (("D", lo), ("E", mid), ("F", hi)):
        put(sc, f"{col}{r}", val, color=BLUE, fmt=fmt, align="center", border=BOX, fill=YEL_FILL)
    put(sc, f"H{r}", why, size=9, color="595959", wrap=True)
    sc.row_dimensions[r].height = 26
    r += 1

put(sc, f"B{r+1}", "Active scenario", bold=True, size=10)
put(sc, f"C{r+1}", "=Inputs!C6", color=GREEN, align="center", bold=True)
put(sc, f"D{r+1}", '=INDEX($D$5:$F$5,1,Inputs!C6)', color=GREEN, bold=True)
put(sc, f"B{r+3}", "Everything not listed here is on the Inputs sheet and applies to all three scenarios.", italic=True, size=9, color="808080")

wb.save(OUT)
print("part 1 saved:", OUT)
