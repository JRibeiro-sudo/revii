# -*- coding: utf-8 -*-
"""Part 6: two dashboard charts. Palette validated (#2a78d6 / #eb6834, all six checks PASS)."""
import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference, Series
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.chart.axis import ChartLines
from openpyxl.styles import Font, PatternFill

OUT="/home/user/revii/financials/Revii_VoiceCare_Financial_Model.xlsx"
S1="2A78D6"   # categorical slot 1
S2="EB6834"   # categorical slot 2
GRID="E8E8E4"; INK="52514E"

wb=openpyxl.load_workbook(OUT)
d=wb["Dashboard"]; a=wb["Annual"]

def label(ref,text):
    d[ref]=text
    d[ref].font=Font(name="Arial",bold=True,size=10,color="1F3864")

def recessive(ch):
    ln=LineProperties(solidFill=GRID,w=6350)
    ch.x_axis.majorGridlines=None
    ch.y_axis.majorGridlines=ChartLines(spPr=GraphicalProperties(ln=ln))
    for ax in (ch.x_axis,ch.y_axis):
        ax.spPr=GraphicalProperties(ln=LineProperties(solidFill=GRID,w=6350))
        ax.txPr=None
        ax.delete=False
    ch.y_axis.numFmt='#,##0'
    ch.legend.position="b"
    ch.legend.overlay=False
    ch.graphical_properties=GraphicalProperties(ln=LineProperties(noFill=True))

# ---- Chart 1: annual net revenue vs EBITDA (same unit, one axis)
label("B72","NET REVENUE AND EBITDA BY YEAR  (EUR)")
c1=BarChart(); c1.type="col"; c1.grouping="clustered"; c1.gapWidth=60; c1.overlap=-10
data=Reference(d,min_col=2,max_col=7,min_row=55,max_row=56)
cats=Reference(d,min_col=3,max_col=7,min_row=54,max_row=54)
c1.add_data(data,titles_from_data=True,from_rows=True)
c1.set_categories(cats)
for s,col in zip(c1.series,(S1,S2)):
    s.graphicalProperties=GraphicalProperties(solidFill=col,
                                              ln=LineProperties(solidFill="FCFCFB",w=25400))
c1.height=8.5; c1.width=19
recessive(c1)
d.add_chart(c1,"B73")

# ---- Chart 2: cumulative equity in vs cumulative free cash flow, monthly
label("B93","CUMULATIVE FOUNDERS' EQUITY IN vs CUMULATIVE FREE CASH FLOW  (EUR, by month)")
c2=LineChart()
m=wb["P&L Monthly"]
eq=Reference(m,min_col=2,max_col=63,min_row=63,max_row=63)
fcf=Reference(m,min_col=2,max_col=63,min_row=60,max_row=60)
c2.add_data(eq,titles_from_data=True,from_rows=True)
c2.add_data(fcf,titles_from_data=True,from_rows=True)
c2.set_categories(Reference(m,min_col=4,max_col=63,min_row=4,max_row=4))
for s,col in zip(c2.series,(S1,S2)):
    s.graphicalProperties=GraphicalProperties(ln=LineProperties(solidFill=col,w=19050))
    s.smooth=False
    s.marker=None
c2.height=8.5; c2.width=19
recessive(c2)
c2.x_axis.tickLblSkip=6
c2.x_axis.tickMarkSkip=6
d.add_chart(c2,"B94")

d["B92"]=("Both series are euros on one axis. The blue line is money the founders have put in; "
          "the orange line is what the business has generated. Where orange crosses blue, the venture has repaid itself.")
d["B92"].font=Font(name="Arial",italic=True,size=9,color="595959")
d["B71"]=("The bars share one axis because both series are euros. Year 1 is short — it contains the build months.")
d["B71"].font=Font(name="Arial",italic=True,size=9,color="595959")

wb.save(OUT)
print("part 6 saved — charts added")
