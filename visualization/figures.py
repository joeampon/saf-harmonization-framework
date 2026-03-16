"""visualization/figures.py — SAF Harmonization Publication Figures (16 figures)."""

import numpy as np
import pandas as pd
import matplotlib; matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch
from matplotlib.gridspec import GridSpec
from matplotlib.lines import Line2D
from scipy.stats import gaussian_kde
from scipy import stats
import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import (PETROLEUM_JET_GHG_WTW, CORSIA_TIER2_THRESHOLD,
                    EU_RED3_THRESHOLD, FIGURES_DIR)
from data.parameter_distributions import METHODOLOGICAL_PARAMS

plt.rcParams.update({"font.family":"DejaVu Sans","font.size":10,
    "axes.linewidth":1.1,"pdf.fonttype":42,"ps.fonttype":42,"figure.dpi":150})

PW  = {"ATJ":"#1565C0","HEFA":"#2E7D32","FT-SPK":"#E65100","PtL":"#6A1B9A"}
PWL = {"ATJ":"#90CAF9","HEFA":"#81C784","FT-SPK":"#FFCC80","PtL":"#CE93D8"}
AM  = {"energy":"o","mass":"s","economic":"^","system_expansion":"D"}
CT  = "#1565C0"; CM = "#C0392B"
PWO = ["ATJ","HEFA","FT-SPK","PtL"]

def _save(fig, name):
    fig.savefig(os.path.join(FIGURES_DIR,f"{name}.png"),dpi=300,bbox_inches="tight",facecolor="white")
    fig.savefig(os.path.join(FIGURES_DIR,f"{name}.pdf"),bbox_inches="tight",facecolor="white")
    plt.close(fig); print(f"  Saved: {name}.{{png,pdf}}")

def _modal(pd_):
    return {n:(s[2] if s[0]=="triangular" else s[1] if s[0]=="normal" else (s[1]+s[2])/2)
            for n,s in pd_.items()}

def _pct(spec,q):
    k=spec[0]
    if k=="triangular":
        _,lo,mo,hi=spec
        if hi<=lo: return mo
        return float(stats.triang.ppf(q,c=(mo-lo)/(hi-lo),loc=lo,scale=hi-lo))
    if k=="uniform":  return spec[1]+q*(spec[2]-spec[1])
    if k=="normal":   return float(stats.norm.ppf(q,loc=spec[1],scale=spec[2]))
    return spec[1]

def _ax_style(ax):
    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False)
    ax.grid(axis="y",alpha=0.25,linestyle=":")

# ── compute modal cost/GHG breakdown ─────────────────────────────────────────
def compute_modal_breakdown():
    from data.parameter_distributions import ATJ_PARAMS,HEFA_PARAMS,FTSPK_PARAMS,PTL_PARAMS
    from config import (JET_LHV_BTU_PER_GAL,GASOLINE_LHV_BTU_PER_GAL,JET_LHV_MJ_PER_L,
                        JET_DENSITY_KG_PER_L,L_JET_PER_GGE,NG_COMBUSTION_GCO2E_MJ,
                        GREY_H2_GCO2E_KG,H2_LHV_MJ_PER_KG,ELEC_KWH_PER_KG_H2,CO2_KG_PER_KG_H2)
    from harmonization.engine import crf,HARMONIZED_PLANT_LIFETIME
    r={}
    p=_modal(ATJ_PARAMS)
    af=2000*365*p["capacity_factor"]; ae=af*p["ethanol_yield"]
    aj=ae*p["jet_yield"]; ad=ae*0.15; ag=ae*0.14
    jg=aj*(JET_LHV_BTU_PER_GAL/GASOLINE_LHV_BTU_PER_GAL)
    dg=ad*(129488/GASOLINE_LHV_BTU_PER_GAL); tg=jg+dg+ag
    cv=crf(p["discount_rate"],HARMONIZED_PLANT_LIFETIME)
    r["ATJ"]={"cost":{"Capital":p["capex_2023"]*cv/tg,"O&M":p["capex_2023"]*0.07/tg,
                       "Feedstock":af*p["feedstock_cost"]/tg},
               "ghg":{"Feedstock GHG":p["feedstock_ghg"]*p["alloc_factor"],
                       "NG combustion":p["ng_use"]*NG_COMBUSTION_GCO2E_MJ*p["alloc_factor"],
                       "Electricity GHG":p["elec_use"]*p["grid_intensity"]*p["alloc_factor"],
                       "Boundary offset":p["boundary_offset"]}}
    p=_modal(HEFA_PARAMS)
    af=800*365*p["capacity_factor"]; ajkg=af*1000*p["jet_yield"]
    ajl=ajkg/JET_DENSITY_KG_PER_L; ajmj=ajl*JET_LHV_MJ_PER_L; ajg=ajl/L_JET_PER_GGE
    cv=crf(p["discount_rate"],HARMONIZED_PLANT_LIFETIME); h2t=af*p["h2_use"]
    h2g=(h2t*1000*GREY_H2_GCO2E_KG)/ajmj
    r["HEFA"]={"cost":{"Capital":p["capex_2023"]*cv/ajg,"O&M":p["capex_2023"]*0.07/ajg,
                        "Feedstock":af*p["feedstock_cost"]/ajg,"H\u2082":h2t*1000*p["h2_price"]/ajg},
                "ghg":{"Feedstock GHG":p["feedstock_ghg"]*p["alloc_factor"],
                        "H\u2082 production":h2g*p["alloc_factor"],
                        "Process GHG":p["process_ghg"]*p["alloc_factor"],
                        "Boundary offset":p["boundary_offset"]}}
    p=_modal(FTSPK_PARAMS)
    af=2500*365*p["capacity_factor"]; jmj=af*17500*p["ft_efficiency"]*0.65
    jl=jmj/JET_LHV_MJ_PER_L; jg=jl/L_JET_PER_GGE
    cv=crf(p["discount_rate"],HARMONIZED_PLANT_LIFETIME)
    r["FT-SPK"]={"cost":{"Capital":p["capex_2023"]*cv/jg,"O&M":p["capex_2023"]*0.05/jg,
                          "Feedstock":af*p["feedstock_cost"]/jg},
                  "ghg":{"Feedstock GHG":p["feedstock_ghg"]*p["alloc_factor"],
                          "Process GHG":p["process_ghg"]*p["alloc_factor"],
                          "Boundary offset":p["boundary_offset"]}}
    p=_modal(PTL_PARAMS)
    MW=200; amwh=MW*8760*p["capacity_factor"]
    akgh2=amwh*1000/ELEC_KWH_PER_KG_H2; aco2t=akgh2*CO2_KG_PER_KG_H2/1000
    jmj=akgh2*H2_LHV_MJ_PER_KG*p["ft_efficiency"]*0.70
    jl=jmj/JET_LHV_MJ_PER_L; jg=jl/L_JET_PER_GGE
    cv=crf(p["discount_rate"],HARMONIZED_PLANT_LIFETIME); ec=p["elec_capex_kw"]*MW*1000
    r["PtL"]={"cost":{"Electrolyzer cap.":ec*cv/jg,"FT capital":p["ft_capex"]*cv/jg,
                       "Electricity":amwh*p["elec_cost_mwh"]/jg,
                       "CO\u2082 capture":aco2t*p["co2_capture_cost"]/jg,
                       "O&M":(ec+p["ft_capex"])*0.04/jg},
               "ghg":{"Electricity GHG":amwh*1000*p["grid_intensity"]/jmj,
                       "Boundary offset":p["boundary_offset"]}}
    return r

def compute_oat_swings(pathway, metric="mfsp"):
    from data.parameter_distributions import PATHWAY_PARAMS
    from models.pathway_models import PATHWAY_MODELS
    pd_=PATHWAY_PARAMS[pathway]; fn=PATHWAY_MODELS[pathway]
    mo=_modal(pd_); mv=fn(mo)[metric]
    sw=[]
    for n,sp in pd_.items():
        try:
            lo=fn({**mo,n:_pct(sp,0.05)})[metric]
            hi=fn({**mo,n:_pct(sp,0.95)})[metric]
        except: lo=hi=mv
        sw.append((n,lo-mv,hi-mv))
    sw.sort(key=lambda x:abs(x[2]-x[1]),reverse=True)
    return sw,mv

# ── Figure functions (condensed but complete) ─────────────────────────────────

def fig1_literature_overview(dh):
    fig,ax=plt.subplots(1,2,figsize=(14,6))
    for a,mc,gc,t in [(ax[0],"mfsp_2023_raw","ghg_raw","(a) Raw"),
                       (ax[1],"mfsp_harmonized","ghg_harmonized","(b) Harmonized")]:
        for pw,g in dh.groupby("pathway"):
            for al,s in g.groupby("allocation"):
                a.scatter(s[mc],s[gc],c=PW.get(pw,"gray"),marker=AM.get(al,"x"),
                          s=75,alpha=0.85,edgecolors="k",linewidths=0.4,zorder=3)
        a.axhline(PETROLEUM_JET_GHG_WTW,color="red",ls="--",lw=1.2,alpha=0.8)
        a.set_xlabel("MFSP (2023 $/GGE)",fontsize=11,fontweight="bold")
        a.set_ylabel("GHG (gCO\u2082e/MJ)",fontsize=11,fontweight="bold")
        a.set_title(t,fontsize=11,fontweight="bold",loc="left")
        a.grid(alpha=0.25,linestyle=":"); a.set_ylim(-15,95); _ax_style(a)
    ax[1].legend(handles=[mpatches.Patch(color=c,label=p) for p,c in PW.items()]+
                 [Line2D([0],[0],marker=m,color="gray",ls="None",ms=8,
                          label=a.replace("_"," ").title()) for a,m in AM.items()],
                 fontsize=8,ncol=2)
    plt.tight_layout(); _save(fig,"Fig1_Literature_Overview")

def fig2_harmonization_impact(dh):
    fig,ax=plt.subplots(1,2,figsize=(14,5))
    for a,rc,hc,yl,t in [(ax[0],"mfsp_2023_raw","mfsp_harmonized","MFSP (2023 $/GGE)","(a) Cost"),
                          (ax[1],"ghg_raw","ghg_harmonized","GHG (gCO\u2082e/MJ)","(b) GHG")]:
        x=np.arange(len(PWO)); w=0.35
        mr=[dh[dh.pathway==p][rc].mean() for p in PWO]; sr=[dh[dh.pathway==p][rc].std() for p in PWO]
        mh=[dh[dh.pathway==p][hc].mean() for p in PWO]; sh=[dh[dh.pathway==p][hc].std() for p in PWO]
        cl=[PW[p] for p in PWO]
        a.bar(x-w/2,mr,w,yerr=sr,color=cl,alpha=0.40,edgecolor="k",capsize=4,label="Raw")
        a.bar(x+w/2,mh,w,yerr=sh,color=cl,alpha=0.90,edgecolor="k",capsize=4,label="Harmonized")
        a.set_xticks(x); a.set_xticklabels(PWO,fontsize=10)
        a.set_ylabel(yl,fontsize=11,fontweight="bold")
        a.set_title(t,fontsize=11,fontweight="bold",loc="left")
        a.legend(fontsize=9); _ax_style(a)
        if "GHG" in t: a.axhline(PETROLEUM_JET_GHG_WTW,color="red",ls="--",lw=1.0)
    plt.tight_layout(); _save(fig,"Fig2_Harmonization_Impact")

def fig3_monte_carlo_distributions(mc):
    fig,ax=plt.subplots(1,2,figsize=(14,6))
    for a,col,xl,t in [(ax[0],"mfsp","MFSP (2023 $/GGE)","(a) MFSP Distribution"),
                        (ax[1],"ghg","GHG (gCO\u2082e/MJ)","(b) GHG Distribution")]:
        data=[mc[p][col].values for p in PWO]; pos=range(1,5)
        vp=a.violinplot(data,positions=pos,showmedians=True,showextrema=False,widths=0.7)
        for body,pw in zip(vp["bodies"],PWO): body.set_facecolor(PWL[pw]); body.set_alpha(0.7)
        vp["cmedians"].set_color("black"); vp["cmedians"].set_linewidth(2)
        for d,p in zip(data,pos):
            q5,q95=np.nanpercentile(d,[5,95])
            a.plot([p,p],[q5,q95],"k-",lw=1.5,zorder=3)
            a.plot([p-0.12,p+0.12],[q5,q5],"k-",lw=1.5); a.plot([p-0.12,p+0.12],[q95,q95],"k-",lw=1.5)
        a.set_xticks(list(pos)); a.set_xticklabels([f"$\\bf{{{p}}}$" for p in PWO],fontsize=11)
        a.set_ylabel(xl,fontsize=11,fontweight="bold"); a.set_title(t,fontsize=11,fontweight="bold",loc="left")
        a.grid(axis="y",alpha=0.25,linestyle=":"); _ax_style(a)
        if col=="ghg": a.axhline(PETROLEUM_JET_GHG_WTW,color="red",ls="--",lw=1.2)
    plt.tight_layout(); _save(fig,"Fig3_MC_Distributions")

def fig4_variance_decomposition(vdf,sr):
    fig=plt.figure(figsize=(18,6)); gs=GridSpec(1,3,figure=fig,wspace=0.38)
    ax1=fig.add_subplot(gs[0]); pw="ATJ"
    s1m=sr[pw]["S1_mfsp"]; s1g=sr[pw]["S1_ghg"]; pars=list(s1m.keys())
    x=np.arange(len(pars)); w=0.35
    cl=[CM if p in METHODOLOGICAL_PARAMS else CT for p in pars]
    ax1.barh(x+w/2,[s1m[p] for p in pars],w,color=cl,alpha=0.85,edgecolor="w",lw=0.5)
    ax1.barh(x-w/2,[s1g[p] for p in pars],w,color=cl,alpha=0.45,edgecolor="w",lw=0.5,hatch="//")
    ax1.set_yticks(x); ax1.set_yticklabels([p.replace("_","\n") for p in pars],fontsize=8)
    ax1.set_xlabel("First-Order Sobol Index S\u2081",fontsize=10,fontweight="bold")
    ax1.set_title(f"(a) Sensitivity Indices\n(ATJ)",fontsize=10,fontweight="bold",loc="left")
    ax1.axvline(0,color="k",lw=0.8); ax1.grid(axis="x",alpha=0.25,linestyle=":")
    ax1.legend(handles=[mpatches.Patch(color=CM,label="Methodological"),
                         mpatches.Patch(color=CT,label="Technical")],fontsize=8,loc="lower right")
    ax2=fig.add_subplot(gs[1]); x2=np.arange(len(PWO)); w2=0.35
    for mi,(metric,ht) in enumerate([("MFSP",""),("GHG","///")]):
        sub=vdf[vdf.Metric==metric]
        mth=[sub[sub.Pathway==p]["Methodological_pct"].values[0] for p in PWO]
        tch=[sub[sub.Pathway==p]["Technical_pct"].values[0] for p in PWO]
        off=x2-w2/2+mi*w2
        ax2.bar(off,mth,w2,color=CM,alpha=0.85,hatch=ht,edgecolor="w",lw=0.5)
        ax2.bar(off,tch,w2,bottom=mth,color=CT,alpha=0.85,hatch=ht,edgecolor="w",lw=0.5)
    ax2.set_xticks(x2); ax2.set_xticklabels(PWO,fontsize=10)
    ax2.set_ylabel("Share of Total Variance (%)",fontsize=10,fontweight="bold")
    ax2.set_title("(b) Variance Decomposition\n(solid=MFSP, hatched=GHG)",fontsize=10,fontweight="bold",loc="left")
    ax2.set_ylim(0,120); ax2.grid(axis="y",alpha=0.25,linestyle=":")
    ax2.legend(handles=[mpatches.Patch(color=CM,label="Methodological"),
                         mpatches.Patch(color=CT,label="Technical")],fontsize=9)
    ax3=fig.add_subplot(gs[2])
    for i,pw in enumerate(PWO):
        for metric,col in [("MFSP","#1565C0"),("GHG","#e53935")]:
            row=vdf[(vdf.Pathway==pw)&(vdf.Metric==metric)].iloc[0]
            cvb,cva=row["CV_before_pct"],row["CV_after_pct"]
            xoff=i+(0.15 if metric=="GHG" else -0.15)
            ax3.annotate("",xy=(xoff,cva),xytext=(xoff,cvb),
                         arrowprops=dict(arrowstyle="->",color=col,lw=2.0))
            ax3.scatter(xoff,cvb,color=col,s=60,marker="o",zorder=5)
            ax3.scatter(xoff,cva,color=col,s=60,marker="D",zorder=5)
    ax3.set_xticks(range(len(PWO))); ax3.set_xticklabels(PWO,fontsize=10)
    ax3.set_ylabel("CV (%)",fontsize=10,fontweight="bold")
    ax3.set_title("(c) CV Before (○) → After (◆)",fontsize=10,fontweight="bold",loc="left")
    ax3.grid(axis="y",alpha=0.25,linestyle=":")
    ax3.legend(handles=[mpatches.Patch(color="#1565C0",label="MFSP"),
                         mpatches.Patch(color="#e53935",label="GHG")],fontsize=9)
    plt.tight_layout(); _save(fig,"Fig4_Variance_Decomposition")

def figA_system_boundary():
    fig,ax=plt.subplots(figsize=(14,5)); ax.set_xlim(0,14); ax.set_ylim(0,5); ax.axis("off")
    stages=[(0.3,2.5,1.8,1.4,"#BBDEFB","Feedstock\nProduction","Well/Farm\nForest/DAC"),
            (2.6,2.5,1.8,1.4,"#C8E6C9","Feedstock\nTransport","Logistics &\nStorage"),
            (4.9,2.5,1.8,1.4,"#FFE0B2","SAF\nConversion","Biorefinery/\nElectrolyzer"),
            (7.2,2.5,1.8,1.4,"#F3E5F5","Fuel\nDistribution","Pipeline/\nTanker"),
            (9.5,2.5,1.8,1.4,"#FFCDD2","Combustion\n(Aircraft)","Jet engine\nexhaust")]
    br=[]
    for x,y,w,h,color,title,sub in stages:
        ax.add_patch(FancyBboxPatch((x,y-h/2),w,h,boxstyle="round,pad=0.07",
                     facecolor=color,edgecolor="#455A64",linewidth=1.5))
        ax.text(x+w/2,y+0.20,title,ha="center",va="center",fontsize=9,fontweight="bold",color="#1A237E")
        ax.text(x+w/2,y-0.28,sub,ha="center",va="center",fontsize=7.5,color="#37474F")
        br.append((x+w,y))
    for i in range(len(br)-1):
        ax.annotate("",xy=(stages[i+1][0],br[i][1]),xytext=(br[i][0],br[i][1]),
                    arrowprops=dict(arrowstyle="->",color="#455A64",lw=1.8,mutation_scale=14))
    for label,x0,x1,color,yoff in [("Well-to-Gate (WtG)",0.3,8.8,"#1565C0",0.0),
                                     ("Well-to-Wake (WtWake)",0.3,11.7,"#B71C1C",-0.45)]:
        yb=1.35+yoff
        ax.annotate("",xy=(x1,yb),xytext=(x0,yb),
                    arrowprops=dict(arrowstyle="<->",color=color,lw=2,mutation_scale=12))
        ax.text((x0+x1)/2,yb-0.27,label,ha="center",fontsize=9,fontweight="bold",color=color)
    ax.set_title("Figure A — Well-to-Wake System Boundary | Functional Unit: 1 MJ SAF",
                 fontsize=10,fontweight="bold",loc="left",pad=8)
    plt.tight_layout(); _save(fig,"FigA_System_Boundary")

def figB_harmonization_flowchart(dh):
    fig,ax=plt.subplots(figsize=(13,9)); ax.set_xlim(0,13); ax.set_ylim(0,10); ax.axis("off")
    steps=[("1","Raw Study Data","As reported: mixed currencies, boundaries, allocation methods","#ECEFF1","#37474F"),
           ("2","CEPCI/CPI Cost Escalation","Capital → 2023 USD (CEPCI₂₀₂₃=798); CPI for opex/feedstock","#E3F2FD","#1565C0"),
           ("3","Allocation Re-weighting","Energy allocation per ISO 14044 §4.3.4.2; pathway-specific factors","#E8F5E9","#1B5E20"),
           ("4","System Boundary Offset","All studies → Well-to-Wake; WtG→WtWake delta (+3.0 gCO₂e/MJ)","#FFF3E0","#E65100"),
           ("5","CRF Normalization","Pathway-specific CAPEX fraction; 10% DR, 30-yr, 90% CF","#F3E5F5","#4A148C"),
           ("\u2713","Harmonized Dataset (42 studies)","WtWake | 1 MJ FU | Energy alloc | 2023 USD | 10% DR | 90% CF","#C8E6C9","#1B5E20")]
    yp=[9.1,7.7,6.3,4.9,3.5,1.9]; bx,bw,bh=1.5,8.5,0.95
    for (num,title,desc,bg,fg),y in zip(steps,yp):
        ax.add_patch(FancyBboxPatch((bx,y-bh/2),bw,bh,boxstyle="round,pad=0.1",
                     facecolor=bg,edgecolor=fg,linewidth=1.8))
        ax.add_patch(plt.Circle((bx+0.45,y),0.28,color=fg,zorder=3))
        ax.text(bx+0.45,y,num,ha="center",va="center",fontsize=9,fontweight="bold",color="white",zorder=4)
        ax.text(bx+1.0,y+0.19,title,ha="left",va="center",fontsize=10,fontweight="bold",color=fg)
        ax.text(bx+1.0,y-0.22,desc,ha="left",va="center",fontsize=8,color="#37474F")
        if y>yp[-1]:
            ax.annotate("",xy=(bx+bw/2,y-bh/2-0.18),xytext=(bx+bw/2,y-bh/2),
                        arrowprops=dict(arrowstyle="->",color="#607D8B",lw=1.8))
    md=(dh["mfsp_harmonized"]-dh["mfsp_2023_raw"]).abs().mean()
    gd=(dh["ghg_harmonized"]-dh["ghg_raw"]).abs().mean()
    ax.text(0.2,0.5,f"Mean |\u0394MFSP|=${md:.2f}/GGE\nMean |\u0394GHG|={gd:.1f} gCO\u2082e/MJ",
            fontsize=8,color="#424242",bbox=dict(boxstyle="round",facecolor="white",edgecolor="#BDBDBD"))
    ax.set_title("Figure B — Five-Step Harmonization Protocol",fontsize=11,fontweight="bold",loc="left",pad=6)
    plt.tight_layout(); _save(fig,"FigB_Harmonization_Flowchart")

def figC_cost_breakdown(bd):
    fig,ax=plt.subplots(figsize=(11,6))
    items=list(dict.fromkeys(k for pw in PWO for k in bd[pw]["cost"]))
    ic={"Capital":"#1565C0","O&M":"#64B5F6","Feedstock":"#388E3C","H\u2082":"#81C784",
        "Electrolyzer cap.":"#7B1FA2","FT capital":"#CE93D8","Electricity":"#F57C00","CO\u2082 capture":"#FDD835"}
    x=np.arange(len(PWO)); bot=np.zeros(len(PWO))
    for item in items:
        vals=np.array([bd[pw]["cost"].get(item,0) for pw in PWO])
        ax.bar(x,vals,bottom=bot,color=ic.get(item,"#9E9E9E"),edgecolor="white",linewidth=0.6,label=item,width=0.55)
        for xi,(v,b) in enumerate(zip(vals,bot)):
            if v>0.3: ax.text(xi,b+v/2,f"{v:.1f}",ha="center",va="center",fontsize=8,color="white",fontweight="bold")
        bot+=vals
    ax.set_xticks(x); ax.set_xticklabels([f"$\\bf{{{pw}}}$" for pw in PWO],fontsize=12)
    for tick,pw in zip(ax.get_xticklabels(),PWO): tick.set_color(PW[pw])
    ax.set_ylabel("MFSP (2023 $/GGE)",fontsize=11,fontweight="bold")
    ax.set_title("Figure C — MFSP Cost Breakdown (Modal Parameter Values)",fontsize=10,fontweight="bold",loc="left")
    ax.set_ylim(0,max(bot)*1.20); ax.legend(loc="upper left",fontsize=8.5,ncol=2,framealpha=0.9); _ax_style(ax)
    plt.tight_layout(); _save(fig,"FigC_Cost_Breakdown")

def figD_tornado_sensitivity():
    fig,axes=plt.subplots(2,2,figsize=(16,12))
    fig.suptitle("Figure D — Tornado Sensitivity (OAT P5→P95)\nRed=methodological | Blue=technical",
                 fontsize=11,fontweight="bold")
    for ax,pw in zip(axes.flat,PWO):
        sw,mv=compute_oat_swings(pw,"mfsp"); top=sw[:8]
        for i,(name,lo,hi) in enumerate(top):
            is_m=name in METHODOLOGICAL_PARAMS; cl=CM if is_m else CT; lg="#FFCDD2" if is_m else "#BBDEFB"
            ax.barh(i,hi,color=cl,alpha=0.78,height=0.55)
            ax.barh(i,lo,color=lg,alpha=0.90,height=0.55,edgecolor=cl,linewidth=0.8)
        ax.axvline(0,color="black",lw=1.2)
        ax.set_yticks(range(len(top))); ax.set_yticklabels([s[0].replace("_"," ") for s in top],fontsize=8.5)
        for yt,(n,_,_) in zip(ax.get_yticklabels(),top): yt.set_color(CM if n in METHODOLOGICAL_PARAMS else CT)
        ax.set_xlabel("\u0394MFSP from modal ($/GGE)",fontsize=9,fontweight="bold")
        ax.set_title(f"{pw}  (modal=${mv:.2f}/GGE)",fontsize=10,fontweight="bold",color=PW[pw])
        ax.grid(axis="x",alpha=0.25,linestyle=":"); _ax_style(ax)
    plt.tight_layout(); _save(fig,"FigD_Tornado_Sensitivity")

def figE_mc_overlay(dh,mc):
    fig,axes=plt.subplots(2,2,figsize=(14,10))
    fig.suptitle("Figure E — Monte Carlo KDE vs Literature",fontsize=11,fontweight="bold")
    for ax,pw in zip(axes.flat,PWO):
        color=PW[pw]; mc_v=mc[pw]["mfsp"].values; mc_v=mc_v[np.isfinite(mc_v)]
        lit=dh[dh.pathway==pw]
        if len(mc_v)>1:
            kde=gaussian_kde(mc_v,bw_method=0.35); xr=np.linspace(mc_v.min(),mc_v.max(),300)
            ax.fill_between(xr,kde(xr),alpha=0.25,color=color); ax.plot(xr,kde(xr),color=color,lw=2,label="MC distribution")
        p5,p50,p95=np.nanpercentile(mc_v,[5,50,95])
        ax.axvline(p50,color=color,ls="--",lw=1.8,label=f"MC P50=${p50:.2f}")
        ax.axvline(p5,color=color,ls=":",lw=1.0,alpha=0.7); ax.axvline(p95,color=color,ls=":",lw=1.0,alpha=0.7)
        ax.autoscale(axis="y"); ax.set_ylim(bottom=-0.015); ytop=ax.get_ylim()[1]
        ax.text(p5,ytop*0.06,"P5",fontsize=7,color=color,ha="center")
        ax.text(p95,ytop*0.06,"P95",fontsize=7,color=color,ha="center")
        if len(lit)>0:
            rv=lit["mfsp_2023_raw"].dropna().values; hv=lit["mfsp_harmonized"].dropna().values
            ax.scatter(rv,np.full(len(rv),-0.005),marker="|",s=120,color="black",alpha=0.8,linewidths=2,
                       label=f"Raw lit (n={len(rv)})",zorder=5)
            ax.scatter(hv,np.full(len(hv),-0.010),marker="|",s=120,color=color,alpha=0.8,linewidths=2,
                       label="Harmonized lit",zorder=5)
        ax.set_xlabel("MFSP (2023 $/GGE)",fontsize=10,fontweight="bold"); ax.set_ylabel("Density",fontsize=9)
        ax.set_title(pw,fontsize=11,fontweight="bold",color=color); ax.legend(fontsize=8)
        ax.grid(alpha=0.20,linestyle=":"); _ax_style(ax)
    plt.tight_layout(); _save(fig,"FigE_MC_Overlay")

def figF_ghg_comparison(mc,dh):
    fig,ax=plt.subplots(figsize=(12,7)); pos=np.arange(1,5)
    data=[mc[p]["ghg"].values for p in PWO]
    vp=ax.violinplot(data,positions=pos,showmedians=True,showextrema=False,widths=0.65)
    for body,pw in zip(vp["bodies"],PWO): body.set_facecolor(PW[pw]); body.set_alpha(0.60)
    vp["cmedians"].set_color("black"); vp["cmedians"].set_linewidth(2.5)
    for d,p in zip(data,pos):
        q5,q25,q75,q95=np.nanpercentile(d,[5,25,75,95])
        ax.plot([p,p],[q5,q95],"k-",lw=1.8,zorder=3)
        ax.fill_betweenx([q25,q75],p-0.09,p+0.09,color="black",alpha=0.20,zorder=4)
    for pw,p in zip(PWO,pos):
        lv=dh[dh.pathway==pw]["ghg_harmonized"].dropna()
        ax.scatter(np.full(len(lv),p),lv,color=PW[pw],s=40,alpha=0.65,zorder=5,edgecolors="k",linewidths=0.5)
    for val,label,color,ls in [(PETROLEUM_JET_GHG_WTW,"Petroleum jet (89)","red","--"),
                                (CORSIA_TIER2_THRESHOLD,"CORSIA -50% (44.5)","#F57C00","-."),
                                (EU_RED3_THRESHOLD,"EU RED III -65% (31.15)","#1565C0",":"),
                                (0.0,"Net-zero","#424242",":")]:
        ax.axhline(val,color=color,ls=ls,lw=1.5,alpha=0.85,label=label)
    ax.set_xticks(list(pos)); ax.set_xticklabels([f"$\\bf{{{pw}}}$" for pw in PWO],fontsize=12)
    for tick,pw in zip(ax.get_xticklabels(),PWO): tick.set_color(PW[pw])
    ax.set_ylabel("GHG (gCO\u2082e/MJ)",fontsize=11,fontweight="bold")
    ax.set_title("Figure F — GHG Distributions vs Regulatory Thresholds",fontsize=10,fontweight="bold",loc="left")
    ax.legend(fontsize=9,loc="upper right"); ax.grid(axis="y",alpha=0.25,linestyle=":"); ax.set_ylim(bottom=-8); _ax_style(ax)
    plt.tight_layout(); _save(fig,"FigF_GHG_Comparison")

def figG_ghg_contribution(bd):
    fig,ax=plt.subplots(figsize=(11,6))
    items=list(dict.fromkeys(k for pw in PWO for k in bd[pw]["ghg"]))
    gc={"Feedstock GHG":"#388E3C","NG combustion":"#F57C00","Electricity GHG":"#FDD835",
        "H\u2082 production":"#0288D1","Process GHG":"#78909C","Boundary offset":"#D32F2F"}
    x=np.arange(len(PWO)); bot=np.zeros(len(PWO))
    for item in items:
        vals=np.array([max(bd[pw]["ghg"].get(item,0),0) for pw in PWO])
        ax.bar(x,vals,bottom=bot,color=gc.get(item,"#9E9E9E"),edgecolor="white",linewidth=0.6,label=item,width=0.55)
        for xi,(v,b) in enumerate(zip(vals,bot)):
            if v>1.0: ax.text(xi,b+v/2,f"{v:.1f}",ha="center",va="center",fontsize=8,color="white",fontweight="bold")
        bot+=vals
    ax.axhline(PETROLEUM_JET_GHG_WTW,color="red",ls="--",lw=1.5,label="Petroleum jet (89)")
    ax.set_xticks(x); ax.set_xticklabels([f"$\\bf{{{pw}}}$" for pw in PWO],fontsize=12)
    for tick,pw in zip(ax.get_xticklabels(),PWO): tick.set_color(PW[pw])
    ax.set_ylabel("GHG (gCO\u2082e/MJ)",fontsize=11,fontweight="bold")
    ax.set_title("Figure G — GHG Component Breakdown (Modal Values; WtWake; Energy Alloc.)",fontsize=10,fontweight="bold",loc="left")
    ax.legend(fontsize=8.5,ncol=2,loc="upper right"); _ax_style(ax)
    plt.tight_layout(); _save(fig,"FigG_GHG_Contribution")

def figH_parameter_heatmap(sr):
    fig,axes=plt.subplots(1,2,figsize=(18,7))
    for ax,(mk,tag,panel) in zip(axes,[("S1_mfsp","MFSP","a"),("S1_ghg","GHG","b")]):
        ap=list(dict.fromkeys(p for pw in PWO for p in sr[pw][mk]))
        mat=pd.DataFrame({pw:[sr[pw][mk].get(p,0) for p in ap] for pw in PWO},index=ap)
        mat=mat.loc[mat.max(axis=1).sort_values(ascending=False).index]
        vmax=min(mat.values.max(),0.70); im=ax.imshow(mat.values,cmap="YlOrRd",aspect="auto",vmin=0,vmax=vmax)
        ax.set_xticks(range(len(PWO))); ax.set_xticklabels(PWO,fontsize=11,fontweight="bold")
        ax.set_yticks(range(len(mat.index))); ax.set_yticklabels([p.replace("_"," ") for p in mat.index],fontsize=9)
        for yt,param in zip(ax.get_yticklabels(),mat.index): yt.set_color(CM if param in METHODOLOGICAL_PARAMS else CT)
        for i in range(mat.shape[0]):
            for j in range(mat.shape[1]):
                v=mat.values[i,j]; ax.text(j,i,f"{v:.2f}",ha="center",va="center",fontsize=8.5,fontweight="bold",
                                            color="white" if v>vmax*0.6 else "black")
        plt.colorbar(im,ax=ax,shrink=0.75,label="S\u2081"); ax.set_title(f"({panel}) {tag} — S\u2081 Heatmap\n(red=methodological | blue=technical)",fontsize=10,fontweight="bold",loc="left")
    fig.suptitle("Figure H — Sobol First-Order Sensitivity Indices",fontsize=11,fontweight="bold")
    plt.tight_layout(); _save(fig,"FigH_Parameter_Heatmap")

def figI_variance_decomposition_extended(vdf,sr):
    fig=plt.figure(figsize=(18,12)); gs=GridSpec(2,3,figure=fig,hspace=0.52,wspace=0.40)
    pi=0
    for row,metric in enumerate(["MFSP","GHG"]):
        mk="S1_mfsp" if metric=="MFSP" else "S1_ghg"; sub=vdf[vdf.Metric==metric]
        ax0=fig.add_subplot(gs[row,0]); x=np.arange(len(PWO))
        mth=[sub[sub.Pathway==p]["Methodological_pct"].values[0] for p in PWO]
        tch=[sub[sub.Pathway==p]["Technical_pct"].values[0] for p in PWO]
        ax0.bar(x,mth,color=CM,alpha=0.85,width=0.6,label="Methodological")
        ax0.bar(x,tch,bottom=mth,color=CT,alpha=0.85,width=0.6,label="Technical")
        for xi,(t,m) in enumerate(zip(tch,mth)):
            if m>5: ax0.text(xi,m/2,f"{m:.0f}%",ha="center",va="center",fontsize=8,fontweight="bold",color="white")
            if t>5: ax0.text(xi,m+t/2,f"{t:.0f}%",ha="center",va="center",fontsize=8,fontweight="bold",color="white")
        ax0.set_xticks(x); ax0.set_xticklabels(PWO,fontsize=9); ax0.set_ylabel("Share (%)",fontsize=9)
        ax0.set_title(f"({chr(97+pi)}) {metric} Variance",fontsize=9,fontweight="bold",loc="left"); ax0.set_ylim(0,118)
        ax0.grid(axis="y",alpha=0.25,linestyle=":"); _ax_style(ax0)
        if row==0: ax0.legend(fontsize=8); pi+=1
        else: pi+=1
        ax1=fig.add_subplot(gs[row,1])
        cvb=[sub[sub.Pathway==p]["CV_before_pct"].values[0] for p in PWO]
        cva=[sub[sub.Pathway==p]["CV_after_pct"].values[0] for p in PWO]
        vr=[sub[sub.Pathway==p]["Variance_Reduction_pct"].values[0] for p in PWO]
        for i,pw in enumerate(PWO):
            ax1.annotate("",xy=(i,cva[i]),xytext=(i,cvb[i]),arrowprops=dict(arrowstyle="->",color=PW[pw],lw=2.2))
            ax1.scatter(i,cvb[i],color=PW[pw],s=70,zorder=5,marker="o"); ax1.scatter(i,cva[i],color=PW[pw],s=70,zorder=5,marker="D")
            ax1.text(i+0.10,(cvb[i]+cva[i])/2,f"-{vr[i]:.0f}%",fontsize=7.5,color=PW[pw])
        ax1.set_xticks(range(len(PWO))); ax1.set_xticklabels(PWO,fontsize=9); ax1.set_ylabel("CV (%)",fontsize=9)
        ax1.set_title(f"({chr(97+pi)}) CV Before → After",fontsize=9,fontweight="bold",loc="left")
        ax1.grid(axis="y",alpha=0.25,linestyle=":"); _ax_style(ax1); pi+=1
        ax2=fig.add_subplot(gs[row,2])
        for j,pw in enumerate(PWO):
            top5=sorted(sr[pw][mk].items(),key=lambda kv:kv[1],reverse=True)[:5]
            for k,(param,val) in enumerate(top5):
                ax2.scatter(val,j+(k-2)*0.14,s=max(val*500+15,15),
                            color=CM if param in METHODOLOGICAL_PARAMS else PW[pw],
                            alpha=0.72,edgecolors="k",linewidths=0.4,zorder=3)
                if val>0.04: ax2.text(val+0.012,j+(k-2)*0.14,param.replace("_"," "),fontsize=6.5,va="center")
        ax2.set_yticks(range(len(PWO))); ax2.set_yticklabels(PWO,fontsize=9)
        ax2.set_xlabel("S\u2081 Index",fontsize=9); ax2.set_title(f"({chr(97+pi)}) Top S\u2081",fontsize=9,fontweight="bold",loc="left")
        ax2.grid(alpha=0.20,linestyle=":"); _ax_style(ax2); pi+=1
    fig.suptitle("Figure I — Extended Variance Decomposition",fontsize=11,fontweight="bold")
    plt.tight_layout(); _save(fig,"FigI_Variance_Decomp_Extended")

def figJ_tea_lca_scatter(dh,mc):
    fig,ax=plt.subplots(figsize=(12,8))
    for pw,g in dh.groupby("pathway"):
        ax.scatter(g["mfsp_harmonized"],g["ghg_harmonized"],c=PW[pw],s=65,alpha=0.80,
                   edgecolors="k",linewidths=0.5,zorder=4,label=f"{pw} (n={len(g)})")
        q90m=g["mfsp_harmonized"].quantile(0.90); q90g=g["ghg_harmonized"].quantile(0.90)
        for _,row in g.iterrows():
            if row["mfsp_harmonized"]>q90m or row["ghg_harmonized"]>q90g:
                ax.annotate(str(row["study_id"])[:12],(row["mfsp_harmonized"],row["ghg_harmonized"]),
                            fontsize=6,alpha=0.75,textcoords="offset points",xytext=(4,3))
    for pw in PWO:
        mm=mc[pw]; med_m=mm["mfsp"].median(); med_g=mm["ghg"].median()
        p5m,p95m=np.nanpercentile(mm["mfsp"],[5,95]); p5g,p95g=np.nanpercentile(mm["ghg"],[5,95])
        ax.errorbar(med_m,med_g,xerr=[[med_m-p5m],[p95m-med_m]],yerr=[[med_g-p5g],[p95g-med_g]],
                    fmt="*",color=PW[pw],ms=16,lw=1.8,capsize=5,markeredgecolor="black",
                    markeredgewidth=0.8,label=f"{pw} MC",zorder=6)
    ax.axhline(PETROLEUM_JET_GHG_WTW,color="red",ls="--",lw=1.5,alpha=0.85,label="Petroleum jet (89)")
    ax.axhline(CORSIA_TIER2_THRESHOLD,color="#F57C00",ls="-.",lw=1.2,alpha=0.85,label="50% GHG cut")
    ax.axhline(0,color="gray",ls=":",lw=0.8)
    ax.set_xlabel("Harmonized MFSP (2023 $/GGE)",fontsize=11,fontweight="bold")
    ax.set_ylabel("Harmonized GHG (gCO\u2082e/MJ)",fontsize=11,fontweight="bold")
    ax.set_title("Figure J — TEA vs LCA Trade-off Space (All 42 Harmonized Studies)",fontsize=10,fontweight="bold",loc="left")
    ax.legend(fontsize=8.5,ncol=2,loc="upper right"); ax.grid(alpha=0.20,linestyle=":"); _ax_style(ax)
    plt.tight_layout(); _save(fig,"FigJ_TEA_LCA_Scatter")

def figK_radar_plot(mc):
    # Cost Competitiveness: scaled from MC median MFSP (lower cost = higher score)
    # GHG Reduction: scaled from MC median GHG vs petroleum baseline
    # Both are computed directly from Monte Carlo results — fully data-driven.
    mfsp50={pw:mc[pw]["mfsp"].median() for pw in PWO}; ghg50={pw:mc[pw]["ghg"].median() for pw in PWO}
    mx=max(mfsp50.values())
    cs={pw:(1-mfsp50[pw]/mx)*5 for pw in PWO}
    gs_={pw:max((PETROLEUM_JET_GHG_WTW-ghg50[pw])/PETROLEUM_JET_GHG_WTW*5,0) for pw in PWO}

    # Technology Readiness Level (TRL) scores — ICAO CORSIA qualified fuel list (2023)
    # and ASTM D4054 qualification status:
    #   HEFA   : TRL 9 (fully commercial, multiple certified plants)
    #   ATJ    : TRL 8 (semi-commercial, Gevo/LanzaTech plants operating)
    #   FT-SPK : TRL 7 (demonstration scale, Red Rock Biofuels)
    #   PtL    : TRL 5 (pilot scale only as of 2023)
    # Scores normalised to 0–5 scale: TRL_score = (TRL / 9) × 5
    trl={"ATJ":8/9*5,"HEFA":5.0,"FT-SPK":7/9*5,"PtL":5/9*5}

    # Land use, water use, feedstock availability: relative expert scores (1=worst, 5=best)
    # Based on comparative assessment in ICAO CORSIA Eligible Fuels LCA (2022) and
    # Prussi et al. (2021) Renewable & Sustainable Energy Reviews DOI 10.1016/j.rser.2021.111614
    #   PtL    : uses no agricultural land (DAC + renewable electricity) → 5
    #   FT-SPK : forestry residues, low land competition → 4
    #   ATJ    : agricultural residues (stover), moderate land use → 3
    #   HEFA   : virgin vegetable oils require dedicated cropland → 2
    land={"ATJ":3.0,"HEFA":2.0,"FT-SPK":4.0,"PtL":5.0}
    water={"ATJ":3.0,"HEFA":2.0,"FT-SPK":3.0,"PtL":4.0}
    feed={"ATJ":4.0,"HEFA":2.0,"FT-SPK":5.0,"PtL":5.0}

    dims=["Cost\nCompetitiveness","GHG\nReduction","Land Use\nEfficiency",
          "Water Use\nEfficiency","TRL\nMaturity","Feedstock\nAvailability"]
    N=len(dims); angles=np.linspace(0,2*np.pi,N,endpoint=False).tolist(); angles+=angles[:1]
    fig,ax=plt.subplots(figsize=(9,9),subplot_kw=dict(polar=True))
    for pw in PWO:
        vals=[cs[pw],gs_[pw],land[pw],water[pw],trl[pw],feed[pw]]; vals+=vals[:1]
        ax.plot(angles,vals,"o-",linewidth=2.2,color=PW[pw],label=pw)
        ax.fill(angles,vals,alpha=0.12,color=PW[pw])
    ax.set_xticks(angles[:-1]); ax.set_xticklabels(dims,fontsize=9,fontweight="bold")
    ax.set_ylim(0,5.6); ax.set_yticks([1,2,3,4,5]); ax.set_yticklabels(["1","2","3","4","5"],fontsize=7,color="gray")
    ax.grid(color="gray",linestyle="--",linewidth=0.5,alpha=0.5)
    ax.set_title(("Figure K — Multi-Dimensional SAF Pathway Comparison\n"
                  "Cost & GHG: Monte Carlo medians  |  TRL: ICAO CORSIA 2023  |  "
                  "Land/Water/Feed: expert scores (Prussi et al. 2021)"),
                 fontsize=8,fontweight="bold",pad=22)
    ax.legend(loc="upper right",bbox_to_anchor=(1.38,1.14),fontsize=10)
    plt.tight_layout(); _save(fig,"FigK_Radar_Plot")

def figL_pareto_sensitivity(sr):
    fig,axes=plt.subplots(2,4,figsize=(22,12),squeeze=False)
    for row,(mk,mn,bc) in enumerate([("S1_mfsp","MFSP",CT),("S1_ghg","GHG","#c62828")]):
        for col,pw in enumerate(PWO):
            ax=axes[row][col]; ax2=ax.twinx()
            s1r=sr[pw][mk]; tot=max(sum(s1r.values()),1e-9)
            ss=sorted(s1r.items(),key=lambda kv:kv[1],reverse=True)
            pars=[kv[0] for kv in ss]; vals=np.array([kv[1] for kv in ss])
            cum=np.cumsum(vals)/tot*100; x=np.arange(len(pars))
            cl=[CM if p in METHODOLOGICAL_PARAMS else bc for p in pars]
            ax.bar(x,vals,color=cl,alpha=0.85,edgecolor="white",linewidth=0.8,width=0.60)
            for xi,v in enumerate(vals):
                if v>0.005: ax.text(xi,v+max(vals)*0.025,f"{v:.2f}",ha="center",fontsize=9,fontweight="bold",
                                     color=CM if pars[xi] in METHODOLOGICAL_PARAMS else "#0d47a1")
            ax2.plot(x,cum,"k-o",ms=5,lw=1.8,zorder=5); ax2.axhline(80,color=CM,ls="--",lw=1.4,alpha=0.85)
            ax2.set_ylim(0,115); ax2.set_ylabel("Cumulative S\u2081 (%)",fontsize=9); ax2.tick_params(axis="y",labelsize=8.5)
            ax2.spines["top"].set_visible(False)
            ax.set_xticks(x); ax.set_xticklabels([p.replace("_"," ") for p in pars],fontsize=8.5,rotation=38,ha="right")
            for xt,param in zip(ax.get_xticklabels(),pars):
                xt.set_color(CM if param in METHODOLOGICAL_PARAMS else "#212121")
                xt.set_fontweight("bold" if param in METHODOLOGICAL_PARAMS else "normal")
            ax.set_ylabel("First-Order Sobol Index S\u2081",fontsize=9)
            ax.set_title(f"{pw} — {mn}",fontsize=11,fontweight="bold",pad=7,color=PW[pw])
            ax.grid(axis="y",alpha=0.28,linestyle=":"); ax.set_ylim(bottom=0); _ax_style(ax); ax.tick_params(axis="y",labelsize=8.5)
            if row==0 and col==0:
                ax.legend(handles=[mpatches.Patch(color=CM,alpha=0.85,label="Methodological"),
                                    mpatches.Patch(color=bc,alpha=0.85,label="Technical")],fontsize=9,loc="upper right")
    plt.tight_layout(pad=2.0,h_pad=3.5,w_pad=2.5); _save(fig,"FigL_Pareto_Sensitivity")

def generate_all_figures(dh,mc,sr,vdf):
    bd=compute_modal_breakdown()
    print("  Fig1"); fig1_literature_overview(dh)
    print("  Fig2"); fig2_harmonization_impact(dh)
    print("  Fig3"); fig3_monte_carlo_distributions(mc)
    print("  Fig4"); fig4_variance_decomposition(vdf,sr)
    print("  FigA"); figA_system_boundary()
    print("  FigB"); figB_harmonization_flowchart(dh)
    print("  FigC"); figC_cost_breakdown(bd)
    print("  FigD"); figD_tornado_sensitivity()
    print("  FigE"); figE_mc_overlay(dh,mc)
    print("  FigF"); figF_ghg_comparison(mc,dh)
    print("  FigG"); figG_ghg_contribution(bd)
    print("  FigH"); figH_parameter_heatmap(sr)
    print("  FigI"); figI_variance_decomposition_extended(vdf,sr)
    print("  FigJ"); figJ_tea_lca_scatter(dh,mc)
    print("  FigK"); figK_radar_plot(mc)
    print("  FigL"); figL_pareto_sensitivity(sr)
