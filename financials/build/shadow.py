# -*- coding: utf-8 -*-
"""Independent re-implementation of the base case, to cross-check the workbook."""
DAYS=30.44; BUILD=4; PT_L=BUILD+1; ES_L=PT_L+12; PREV_L=PT_L+24
PRICE_V=18.99; PRICE_P=8.99; VAT={'pt':.23,'es':.21}; INFL=.01; LEAK=.04; FEE=.025
ADDS0={'pt':10,'es':15}; GROWTH=[.20,.07,.04,.025,.015]; RAMP=1.0; CAP={'pt':500,'es':1000}
CHURN_A=.18; CHURN_M=1-(1-CHURN_A)**(1/12); TRIAL_D=14; TRIAL_C=.35
ATTACH=[0,0,.10,.20,.30]; ATT_M=1.0; PREV_COST=3.60
CALLS=1.0; MINS=2.5; MINS_MO=CALLS*MINS*DAYS; CPM0=.15; CPM_CHG=-.15; DATA=.03
COCA=[38,32,26,22,20]; COCA_M=1.0
OPEX={'found':[0]*5,'dev':[1500,4000,8000,12000,16000],'sup':[0,1200,3000,5000,7000],
      'mkt':[1000,2500,5000,7000,9000],'cloud':[250,400,700,1000,1400],
      'ga':[400,500,700,900,1100],'ins':[0,625,625,750,900],'off':[0,300,600,900,1200],
      'es':[0,1500,3000,4000,5000],'other':[200,400,800,1200,1600]}
MKT_M=1.0
CAPEX_V=46000; CAPEX_P=24250; CAPEX_ES=10000; CAPEX_ONG=.20; DEP_Y=5
RECV_D=5; PAY_D=30; PREPAY=0.0; MINCASH=15000; TAX=.19

netfact={m:(1/(1+VAT[m]))*(1-LEAK)*(1-FEE) for m in ('pt','es')}
cpm=[CPM0*(1+CPM_CHG)**i for i in range(5)]

def market(mk, launch):
    rows=[]; open_=0.0; prev_adds=0.0
    for mth in range(1,61):
        my=0 if mth<launch else (mth-launch)//12+1
        my=min(my,5)
        mm=0 if mth<launch else mth-launch+1
        yr=min(5,(mth-1)//12+1)
        if mm<1: adds=0.0
        elif mm==1: adds=ADDS0[mk]
        else: adds=min(CAP[mk], prev_adds*(1+GROWTH[my-1]*RAMP))
        prev_adds=adds
        churn=open_*CHURN_M
        close=open_+adds-churn
        avg=(open_+close)/2
        trials=adds/TRIAL_C if TRIAL_C else 0
        pv=PRICE_V*(1+INFL)**(yr-1); pp=PRICE_P*(1+INFL)**(yr-1)
        att=0 if (my==0 or mth<PREV_L) else ATTACH[my-1]*ATT_M
        gb_v=avg*pv; gb_p=avg*att*pp
        nr=(gb_v+gb_p)*netfact[mk]
        cogs=avg*MINS_MO*cpm[yr-1]+trials*TRIAL_D*CALLS*MINS*cpm[yr-1]+avg*DATA+avg*att*PREV_COST
        coca=0 if my==0 else COCA[my-1]*COCA_M
        rows.append(dict(mth=mth,yr=yr,open=open_,adds=adds,close=close,avg=avg,
                         gb=gb_v+gb_p,nr=nr,cogs=cogs,acq=adds*coca))
        open_=close
    return rows

pt=market('pt',PT_L); es=market('es',ES_L)
cash=0.0; pool=0.0; cumcap=0.0; accdep=0.0; nwc_prev=0.0; cumeq=0.0; cumfcf=0.0
ann={y:{k:0.0 for k in ('nr','cogs','opex','ebitda','dep','ebit','tax','ni','capex','fcf','eq')} for y in range(1,6)}
eoy={}
for i in range(60):
    mth=i+1; yr=min(5,(mth-1)//12+1); j=yr-1
    nr=pt[i]['nr']+es[i]['nr']; cogs=pt[i]['cogs']+es[i]['cogs']
    gp=nr-cogs
    acq=pt[i]['acq']+es[i]['acq']
    opex=(acq+OPEX['mkt'][j]*MKT_M+OPEX['found'][j]+OPEX['dev'][j]+OPEX['sup'][j]
          +(OPEX['es'][j] if mth>=ES_L else 0)+OPEX['cloud'][j]
          +OPEX['ga'][j]+OPEX['ins'][j]+OPEX['off'][j]+OPEX['other'][j])
    ebitda=gp-opex
    cap=(CAPEX_V/BUILD if mth<=BUILD else 0)+(CAPEX_P if mth==PREV_L else 0)+ \
        (CAPEX_ES if mth==ES_L else 0)+ \
        (0 if yr==1 else (CAPEX_V+(CAPEX_P if mth>PREV_L else 0)+(CAPEX_ES if mth>ES_L else 0))*CAPEX_ONG/12)
    dep=cumcap/(DEP_Y*12)          # opening cumulative capex
    cumcap+=cap; accdep+=dep
    ebit=ebitda-dep
    tax=max(0.0,ebit-pool)*TAX; pool=max(0.0,pool-ebit); ni=ebit-tax
    recv=nr*(1-PREPAY)*RECV_D/DAYS; pay=(cogs+opex)*PAY_D/DAYS
    nwc=recv-pay
    dnwc=nwc-nwc_prev; nwc_prev=nwc
    fcf=ebitda-tax-cap-dnwc
    cumfcf+=fcf
    pre=cash+fcf; inj=max(0.0,MINCASH-pre); cash=pre+inj; cumeq+=inj
    for k,v in (('nr',nr),('cogs',cogs),('opex',opex),('ebitda',ebitda),('dep',dep),
                ('ebit',ebit),('tax',tax),('ni',ni),('capex',cap),('fcf',fcf),('eq',inj)):
        ann[yr][k]+=v
    eoy[yr]=dict(cash=cash,cumeq=cumeq,cumfcf=cumfcf,
                 pt=pt[i]['close'],es=es[i]['close'],nfa=cumcap-accdep)

print(f"{'':22}"+"".join(f"{'Y'+str(y):>13}" for y in range(1,6)))
for k,lab in [('nr','Net revenue'),('cogs','Cost of service'),('ebitda','EBITDA'),
              ('dep','Depreciation'),('tax','Tax'),('ni','Net income'),
              ('capex','Capex'),('fcf','Free cash flow'),('eq','Equity injected')]:
    print(f"{lab:22}"+"".join(f"{ann[y][k]:>13,.0f}" for y in range(1,6)))
for k,lab in [('cash','Cash (year end)'),('cumeq','Cum. equity'),('cumfcf','Cum. FCF'),
              ('pt','Customers PT'),('es','Customers ES')]:
    print(f"{lab:22}"+"".join(f"{eoy[y][k]:>13,.0f}" for y in range(1,6)))
print(f"\nGross margin Y1..Y5: "+", ".join(f"{1-ann[y]['cogs']/ann[y]['nr']:.1%}" for y in range(1,6)))
print(f"Total equity required: EUR {eoy[5]['cumeq']:,.0f}")
print(f"Net rev/customer/mo Y1 PT: EUR {PRICE_V*netfact['pt']:.2f} | AI cost/mo: EUR {MINS_MO*CPM0+DATA:.2f}")
