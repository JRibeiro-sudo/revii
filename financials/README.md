# Revii — Voice Care standalone financial model

`Revii_VoiceCare_Financial_Model.xlsx` models a real launch of **Voice Care only**,
Portugal first and Spain twelve months later, funded 100% by founders' equity.

It is a sizing tool: how much the founders must put in, when the business stops
consuming cash, and what the return looks like. It is not a set of statutory accounts.

## Sheets

| Sheet | What it holds |
|---|---|
| README | How to use it, the colour legend, and the decisions behind the base case |
| Dashboard | Equity required, break-even months, Year 5 scale, unit economics, return |
| Scenarios | The ten levers that swing the business, in three columns |
| Inputs | Every other assumption, each with its source noted alongside |
| Model PT / Model ES | The monthly engine per market — kept separate so the two revenue streams never blend |
| P&L Monthly | 60 months consolidated: P&L, capex, working capital, cash flow, equity injection |
| Annual | Years 1–5 P&L, cash flow and a light balance sheet |

Switch scenario on `Inputs!C7` (1 Conservative / 2 Base / 3 Optimistic). Every blue
cell is editable; black cells are formulas.

## Where the numbers come from

Built from `Financials_Revii_v3_full_product.xlsx` (Voice Care price, COGS build-up,
COCA, churn, CAPEX, OPEX structure), the Lisbon MBA business plan, and the final deck.
Where the Excel and the deck disagree, the deck wins — per the founders' instruction.

Two departures from the MBA model are deliberate and material:

- **VAT.** Neither the plan nor the deck mentions it. The listed price is therefore
  treated as what the customer pays, so net revenue is the price divided by (1 + VAT).
  That removes ~19% of headline revenue in Portugal before any cost.
- **Free trial cost.** The plan relies on a free trial but never costs the AI minutes
  it burns. This model does.

## Rebuilding

`build/build_p1.py` … `build_p6.py` regenerate the workbook in order.
`build/shadow.py` is an independent Python re-implementation of the base case, used
to cross-check the spreadsheet's output rather than trusting a clean recalculation.
