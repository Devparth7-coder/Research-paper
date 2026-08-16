from pathlib import Path
import json, csv, os, math, re
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Circle, FancyArrowPatch, Polygon
from matplotlib import patheffects

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.dml import MSO_THEME_COLOR_INDEX

from pptx import Presentation
from pptx.util import Inches as PInches, Pt as PPt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.dml.color import RGBColor as PRGBColor

ROOT = Path('/home/user/mas_reliab')
FIG = ROOT / 'figures'
SUP = ROOT / 'support'
ROOT.mkdir(exist_ok=True); FIG.mkdir(exist_ok=True); SUP.mkdir(exist_ok=True)

NAVY='#0B1F33'; BLUE='#146C94'; TEAL='#1B998B'; CYAN='#54C6EB'; AMBER='#F4A261'; RED='#D95D5D'; GREEN='#2A9D8F'; LIGHT='#F3F6F8'; MID='#DCE5EA'; DARK='#23323D'; GRAY='#657681'; PURPLE='#7A5AF8'

plt.rcParams.update({'font.family':'DejaVu Sans','font.size':10,'axes.titleweight':'bold'})

def clean_ax(ax):
    ax.set_xlim(0,1); ax.set_ylim(0,1); ax.axis('off')

def box(ax, xy, w, h, text, fc=LIGHT, ec=BLUE, fs=10, lw=1.5, radius=.025, color=DARK, weight='normal'):
    x,y=xy
    p=FancyBboxPatch((x,y),w,h,boxstyle=f"round,pad=0.012,rounding_size={radius}",facecolor=fc,edgecolor=ec,linewidth=lw)
    ax.add_patch(p); ax.text(x+w/2,y+h/2,text,ha='center',va='center',fontsize=fs,color=color,weight=weight,wrap=True)
    return p

def arrow(ax, a, b, color=GRAY, lw=1.7, style='-|>', rad=0):
    ax.add_patch(FancyArrowPatch(a,b,arrowstyle=style,mutation_scale=12,color=color,linewidth=lw,connectionstyle=f'arc3,rad={rad}'))

def savefig(name, fig):
    path=FIG/name
    fig.savefig(path,dpi=260,bbox_inches='tight',facecolor='white')
    plt.close(fig)
    return path

def make_figures():
    # 1 Layered architecture
    fig,ax=plt.subplots(figsize=(11,6.2)); clean_ax(ax)
    ax.text(.5,.96,'MAS-RELIAB as a methodology layer above execution infrastructure',ha='center',va='top',fontsize=16,weight='bold',color=NAVY)
    layers=[
      (.08,.72,.84,.16,'MAS-RELIAB methodology layer','Benchmark catalog  •  fault injector  •  graph builder  •  metric engine  •  statistical analysis',TEAL),
      (.08,.48,.84,.16,'AI Command Center: execution and observability','Episode launcher  •  agent orchestration  •  full trace capture  •  artifact store  •  dashboards',BLUE),
      (.08,.24,.84,.16,'Agent and model layer','Single agent  •  sequential team  •  parallel team  •  hierarchy  •  model/tool adapters',PURPLE),
      (.08,.02,.84,.14,'Controlled task environments','Stateful simulators  •  versioned tools  •  deterministic validators  •  snapshots',AMBER)]
    for x,y,w,h,title,sub,c in layers:
        p=FancyBboxPatch((x,y),w,h,boxstyle='round,pad=.014,rounding_size=.025',facecolor=c+'18',edgecolor=c,linewidth=2)
        ax.add_patch(p); ax.text(x+.03,y+h*.64,title,ha='left',va='center',fontsize=12,weight='bold',color=c)
        ax.text(x+.03,y+h*.30,sub,ha='left',va='center',fontsize=9.5,color=DARK)
    for y in [.71,.47,.23]: arrow(ax,(.5,y),(.5,y-.055),color=GRAY)
    savefig('figure_01_architecture.png',fig)

    # 2 Graph and cascade
    fig,ax=plt.subplots(figsize=(11,5.7)); clean_ax(ax)
    ax.text(.5,.97,'Directed dependency graph, injected origin, and downstream effects',ha='center',va='top',fontsize=16,weight='bold',color=NAVY)
    nodes={'A1':(.12,.62),'A2':(.35,.76),'A3':(.35,.45),'A4':(.62,.72),'A5':(.62,.36),'A6':(.86,.55)}
    edges=[('A1','A2'),('A1','A3'),('A2','A4'),('A3','A4'),('A3','A5'),('A4','A6'),('A5','A6')]
    for u,v in edges: arrow(ax,(nodes[u][0]+.035,nodes[u][1]),(nodes[v][0]-.045,nodes[v][1]),color='#91A4AF',lw=2)
    status={'A1':'ok','A2':'origin','A3':'ok','A4':'affected','A5':'contained','A6':'affected'}
    for n,(x,y) in nodes.items():
        st=status[n]; fc={'ok':TEAL,'origin':RED,'affected':AMBER,'contained':BLUE}[st]
        circ=Circle((x,y),.055,facecolor=fc,edgecolor='white',linewidth=2); ax.add_patch(circ)
        ax.text(x,y,n,ha='center',va='center',color='white',weight='bold',fontsize=11)
        ax.text(x,y-.085,{'ok':'valid','origin':'injected fault','affected':'downstream failure','contained':'detected / contained'}[st],ha='center',va='top',fontsize=8.5,color=DARK)
    box(ax,(.18,.06),.64,.16,'Measurements: event propagation rate • downstream affected fraction • amplification • depth • edge transmission • time to detection',fc=LIGHT,ec=BLUE,fs=10)
    savefig('figure_02_graph.png',fig)

    # 3 Taxonomy
    fig,ax=plt.subplots(figsize=(11,6.3)); clean_ax(ax)
    ax.text(.5,.97,'Operational failure taxonomy used by MAS-RELIAB',ha='center',va='top',fontsize=16,weight='bold',color=NAVY)
    categories=[
      ('Agent-local',RED,['reasoning error','role non-compliance','memory/state error']),
      ('Handoff / edge',AMBER,['data gap','signal corruption','referential drift','capability mismatch']),
      ('Tool / environment',BLUE,['timeout / rate limit','wrong or partial response','schema drift','stale state']),
      ('Coordination',PURPLE,['conflicting plans','duplicate work','deadlock / loop','premature consensus']),
      ('Verification / recovery',TEAL,['missed detection','false alarm','ineffective retry','rollback failure'])]
    xs=[.05,.24,.43,.62,.81]
    for (title,c,items),x in zip(categories,xs):
        box(ax,(x,.72),.15,.12,title,fc=c+'18',ec=c,fs=10,weight='bold',color=c)
        for i,it in enumerate(items):
            box(ax,(x,.56-i*.12),.15,.085,it,fc='white',ec=MID,fs=8.5,lw=1,color=DARK)
            if i<3: arrow(ax,(x+.075,.555-i*.12),(x+.075,.53-i*.12),color=MID,lw=1)
    ax.text(.5,.09,'Each label is operationalized by trigger, payload, target, precondition, expected trace evidence, success check, and seed.',ha='center',va='center',fontsize=10,color=DARK)
    ax.text(.5,.035,'The taxonomy is related to MAST and AgentAsk; it is not claimed as a replacement for either.',ha='center',va='center',fontsize=9,color=GRAY,style='italic')
    savefig('figure_03_taxonomy.png',fig)

    # 4 topologies
    fig,axs=plt.subplots(1,4,figsize=(12,4.3))
    titles=['Single agent','Sequential','Parallel + aggregator','Hierarchical']
    for ax,t in zip(axs,titles): clean_ax(ax); ax.set_title(t,color=NAVY,fontsize=12,pad=12)
    # single
    box(axs[0],(.27,.38),.46,.24,'Agent',fc=TEAL+'18',ec=TEAL,fs=11,weight='bold'); arrow(axs[0],(.08,.50),(.25,.50)); arrow(axs[0],(.74,.50),(.92,.50))
    axs[0].text(.08,.56,'task',ha='center',fontsize=8); axs[0].text(.92,.56,'answer',ha='center',fontsize=8)
    # seq
    for i,x in enumerate([.04,.36,.68]): box(axs[1],(x,.40),.25,.20,f'A{i+1}',fc=BLUE+'18',ec=BLUE,fs=10,weight='bold')
    arrow(axs[1],(.29,.50),(.35,.50)); arrow(axs[1],(.61,.50),(.67,.50))
    # parallel
    for i,y in enumerate([.70,.40,.10]): box(axs[2],(.05,y),.28,.17,f'Worker {i+1}',fc=PURPLE+'18',ec=PURPLE,fs=8.5)
    box(axs[2],(.64,.36),.30,.23,'Aggregator',fc=AMBER+'18',ec=AMBER,fs=9,weight='bold')
    for y in [.785,.485,.185]: arrow(axs[2],(.34,y),(.63,.49),color=GRAY,lw=1.3)
    # hierarchy
    box(axs[3],(.35,.72),.30,.16,'Manager',fc=RED+'18',ec=RED,fs=9,weight='bold')
    for x in [.04,.36,.68]: box(axs[3],(x,.36),.27,.16,'Specialist',fc=BLUE+'18',ec=BLUE,fs=8)
    box(axs[3],(.35,.06),.30,.16,'Verifier',fc=TEAL+'18',ec=TEAL,fs=8.5,weight='bold')
    for x in [.175,.495,.815]: arrow(axs[3],(.5,.71),(x,.53),lw=1.2); arrow(axs[3],(x,.35),(.5,.23),lw=1.2)
    fig.suptitle('Budget-matched architecture baselines',fontsize=16,weight='bold',color=NAVY,y=1.03)
    fig.text(.5,.01,'Primary analysis equalizes task set, seed blocks, model family, token/tool-call caps, and validator.',ha='center',fontsize=9,color=GRAY)
    savefig('figure_04_topologies.png',fig)

    # 5 experimental matrix
    fig,ax=plt.subplots(figsize=(11,6.2)); clean_ax(ax)
    ax.text(.5,.97,'Six-experiment program and controlled factors',ha='center',va='top',fontsize=16,weight='bold',color=NAVY)
    rows=[
      ('E1','Baseline reliability','architecture × task × repeated seed','RQ1 / H6'),
      ('E2','Fault propagation','fault type × position × severity','RQ2 / H1'),
      ('E3','Attribution','trace view × attribution method','RQ3 / H2'),
      ('E4','Verification','none × final-only × local','RQ4 / H3'),
      ('E5','Recovery','retry × rollback × alternate × human gate','RQ5 / H4'),
      ('E6','Topology sensitivity','single × sequential × parallel × hierarchy','RQ6 / H5')]
    for i,(e,name,factors,map_) in enumerate(rows):
        y=.80-i*.125
        box(ax,(.05,y),.08,.075,e,fc=NAVY,ec=NAVY,fs=10,color='white',weight='bold')
        box(ax,(.15,y),.25,.075,name,fc=LIGHT,ec=MID,fs=9.5,weight='bold')
        box(ax,(.42,y),.39,.075,factors,fc='white',ec=MID,fs=9)
        box(ax,(.83,y),.12,.075,map_,fc=TEAL+'18',ec=TEAL,fs=8.5,color=TEAL,weight='bold')
    ax.text(.5,.055,'Common controls: paired task/seed blocks • interleaved scheduling • versioned environment • deterministic injection check • cost and latency logging',ha='center',fontsize=9.5,color=DARK)
    savefig('figure_05_experiments.png',fig)

    # 6 metric stack
    fig,ax=plt.subplots(figsize=(11,6.3)); clean_ax(ax)
    ax.text(.5,.97,'Multidimensional reporting: no scientifically privileged composite score',ha='center',va='top',fontsize=16,weight='bold',color=NAVY)
    metrics=[('Outcome','task success\nend-state validity',TEAL),('Repeatability','pass^k\npairwise state equivalence',BLUE),('Robustness','perturbation and fault\ndegradation',PURPLE),('Propagation','rate • fraction\namplification • depth',RED),('Attribution','agent • step • top-k\nlocalization distance',AMBER),('Resources','tokens • latency\ntool calls • cost',GRAY)]
    for i,(t,s,c) in enumerate(metrics):
        col=i%3; row=i//3; x=.085+col*.305; y=.59-row*.26
        box(ax,(x,y),.255,.19,t+'\n'+s,fc=c+'18',ec=c,fs=9.5,color=c,weight='bold')
    ax.text(.5,.255,'Report each component with uncertainty; use Pareto frontiers for reliability–cost trade-offs.',ha='center',fontsize=10.5,color=DARK,weight='bold')
    box(ax,(.15,.07),.70,.13,'Composite scoring is permitted only as a labelled sensitivity analysis with predeclared weights\nand alternative-weight robustness checks.',fc=LIGHT,ec=RED,fs=8.6,color=DARK)
    savefig('figure_06_metrics.png',fig)

    # 7 observability
    fig,ax=plt.subplots(figsize=(11,5.7)); clean_ax(ax)
    ax.text(.5,.97,'Observability-conditioned failure attribution',ha='center',va='top',fontsize=16,weight='bold',color=NAVY)
    views=[('Output only',['final answer'],.08,RED),('Partial trace',['messages','tool names','final answer'],.365,AMBER),('Full trace',['prompts + inputs','messages + tool I/O','state deltas + timestamps'],.65,TEAL)]
    for title,items,x,c in views:
        box(ax,(x,.62),.25,.12,title,fc=c+'18',ec=c,fs=11,color=c,weight='bold')
        for i,it in enumerate(items): box(ax,(x,.46-i*.11),.25,.075,it,fc='white',ec=MID,fs=8.8)
        box(ax,(x,.10),.25,.10,'Same failed episodes\nSame gold origin labels',fc=LIGHT,ec=BLUE,fs=8.5,color=BLUE)
    arrow(ax,(.205,.80),(.775,.80),color=NAVY,lw=2,style='-|>')
    ax.text(.49,.84,'increasing observable evidence',ha='center',fontsize=9,color=NAVY)
    ax.text(.5,.025,'Primary comparison is paired; any accuracy difference is attributed to the evidence view, not a different failure set.',ha='center',fontsize=9,color=GRAY)
    savefig('figure_07_observability.png',fig)

    # 8 mitigation ladder
    fig,ax=plt.subplots(figsize=(11,6.0)); clean_ax(ax)
    ax.text(.5,.97,'Mitigation ladder and ablation logic',ha='center',va='top',fontsize=16,weight='bold',color=NAVY)
    stages=[('0','No mitigation','reference'),('1','Retry','same agent / context'),('2','Local verification','check before handoff'),('3','Rollback','restore checkpoint'),('4','Alternate agent','independent re-execution'),('5','Human gate','high-risk escalation')]
    colors=[GRAY,BLUE,TEAL,PURPLE,AMBER,RED]
    coords=[(.07,.59),(.38,.59),(.69,.59),(.69,.30),(.38,.30),(.07,.30)]
    for i,((n,t,s),(x,y)) in enumerate(zip(stages,coords)):
        c=colors[i]
        box(ax,(x,y),.235,.16,n+'  '+t+'\n'+s,fc=c+'18',ec=c,fs=8.7,color=c,weight='bold')
    arrow(ax,(.305,.67),(.375,.67),color=GRAY,lw=1.6)
    arrow(ax,(.615,.67),(.685,.67),color=GRAY,lw=1.6)
    arrow(ax,(.807,.58),(.807,.475),color=GRAY,lw=1.6)
    arrow(ax,(.685,.38),(.615,.38),color=GRAY,lw=1.6)
    arrow(ax,(.375,.38),(.305,.38),color=GRAY,lw=1.6)
    ax.text(.5,.12,'Compare recovered task success, residual propagation, false interventions,',ha='center',fontsize=9.5,color=DARK)
    ax.text(.5,.075,'added latency/tokens/cost, and Pareto dominance.',ha='center',fontsize=9.5,color=DARK)
    savefig('figure_08_mitigation.png',fig)

make_figures()
