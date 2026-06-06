"""
TNN Network Analyzer — Connecting for Change
v5 — cleaner UI, projectable insights export
"""

import io, zipfile, xml.etree.ElementTree as ET, math, collections
import pandas as pd
import networkx as nx
import plotly.graph_objects as go
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="TNN Network Analyzer", page_icon="🌐",
                   layout="wide", initial_sidebar_state="expanded")

# ── Palette ───────────────────────────────────────────────────────────────────
INDIGO="#2825BE"; INDIGO_LT="#4a47d6"; AMBER="#EB9001"; TEAL="#0C7A7A"
TEAL_LT="#0fa8a8"; TERRA="#CF4C38"; GREEN="#2E9E5B"; INK="#080818"
INK2="#0f0e2e"; INK3="#13123a"; SURFACE="#ffffff"; SURFACE2="#F5F7F6"
BORDER="#e2e5e3"; TEXT="#111827"; TEXT2="#4b5563"; TEXT3="#9ca3af"

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@600;700;800;900&family=IBM+Plex+Sans:ital,wght@0,300;0,400;0,500;0,600;0,700;1,400&family=IBM+Plex+Mono:wght@400;500&display=swap');
html,body,[class*="css"]{font-family:'IBM Plex Sans',sans-serif;color:#111827;}
h1,h2,h3{font-family:'Barlow Condensed',sans-serif;letter-spacing:-0.01em;color:#080818;}
.eyebrow{font-size:10.5px;font-weight:700;letter-spacing:0.16em;text-transform:uppercase;color:#EB9001;margin-bottom:4px;font-family:'IBM Plex Sans',sans-serif;}
.metric-card{background:#F5F7F6;border-radius:8px;padding:20px 24px;border-left:3px solid #2825BE;}
.metric-card .value{font-family:'Barlow Condensed',sans-serif;font-size:44px;font-weight:800;color:#2825BE;line-height:1;}
.metric-card .label{font-family:'IBM Plex Mono',monospace;font-size:11px;color:#4b5563;margin-top:4px;}
.scorecard-row{display:flex;gap:16px;margin:16px 0;}
.scorecard{flex:1;background:#F5F7F6;border-radius:8px;padding:18px 20px;border-top:3px solid #2825BE;}
.scorecard .sc-val{font-family:'Barlow Condensed',sans-serif;font-size:36px;font-weight:800;line-height:1;margin-bottom:4px;}
.scorecard .sc-label{font-family:'IBM Plex Mono',monospace;font-size:10px;color:#4b5563;text-transform:uppercase;letter-spacing:0.08em;}
        sc_html+=f'''<div class="scorecard" style="border-top-color:{c};">
<div class="sc-val" style="color:{c};">{s}%</div>
<div class="sc-label">{attr}</div>
{"<div class="sc-sub">" + sub + "</div>" if sub else ""}
</div>'''
.finding{border-bottom:1px solid #e2e5e3;padding:20px 0;}
.finding:last-child{border-bottom:none;}
.finding .f-label{font-family:'IBM Plex Mono',monospace;font-size:10px;color:#EB9001;text-transform:uppercase;letter-spacing:0.12em;margin-bottom:6px;}
.finding .f-head{font-family:'Barlow Condensed',sans-serif;font-size:22px;font-weight:800;color:#080818;margin-bottom:6px;line-height:1.2;}
.finding .f-body{font-family:'IBM Plex Sans',sans-serif;font-size:14px;color:#4b5563;line-height:1.7;max-width:680px;}
.finding .f-stat{font-family:'IBM Plex Mono',monospace;font-size:28px;font-weight:500;color:#2825BE;margin-bottom:4px;}
.skill-match{background:#F5F7F6;border-radius:6px;padding:12px 16px;margin-bottom:8px;border-left:2px solid #0C7A7A;}
.skill-match .sm-skill{font-family:'IBM Plex Mono',monospace;font-size:11px;color:#0C7A7A;margin-bottom:4px;text-transform:uppercase;letter-spacing:0.06em;}
.skill-match .sm-body{font-family:'IBM Plex Sans',sans-serif;font-size:13px;color:#111827;}
.stTabs [data-baseweb="tab"]{font-family:'IBM Plex Sans',sans-serif;font-size:13px;font-weight:500;color:#4b5563;}
.stTabs [aria-selected="true"]{color:#2825BE !important;}
</style>""", unsafe_allow_html=True)

# ── Lookup tables ─────────────────────────────────────────────────────────────
FREQUENCY_LABELS={1:"Never or rarely",2:"Once or twice a month",3:"Weekly",4:"Multiple times a week",5:"Daily or near-daily"}
DEPTH_LABELS={1:"Awareness",2:"Connection",3:"Cooperation",4:"Collaboration"}
KNOWLEDGE_LABELS={5:"Not at all",6:"Somewhat",7:"Quite a bit",8:"A great deal"}
SCALE4_LABELS={1:"Not at all",2:"Somewhat",3:"Quite a bit",4:"A great deal"}
TYPE_LABELS={"Type_1":"Peer Learning / Knowledge Exchange","Type_2":"Joint Programming or Co-delivery","Type_3":"Referrals","Type_4":"Funding","Type_5":"Informal Support / Thought Partnership","Type_6":"Governance or Board-level Connection"}
RECIPROCITY_LABELS={1:"Mostly flows from them to me",2:"Roughly balanced",3:"Mostly flows from me to them",4:"I'm not sure"}
SECTOR_LABELS={1:"Education",2:"Health & Human Services",3:"Community Development & Housing",4:"Government & Public Sector",5:"Faith Communities",6:"Business & Commerce",7:"Arts, Culture & Humanities",8:"Civic & Advocacy",9:"Media & Communications",10:"Environment & Conservation",11:"Philanthropy & Funding",12:"Other"}
SHARING_LABELS={1:"Place & Built Environment",2:"Recognition",3:"Rhythm & Recurrence",4:"Mutual Obligation",5:"Common Stakes",6:"Institutional Anchors",7:"Story, Memory & Identity"}
CONFIDENTIALITY_LABELS={1:"Yes – share full responses",2:"Yes – aggregate/anonymized only",3:"No – keep private"}
GML_NS="http://graphml.graphdrawing.org/xmlns"; MAX_BARE=3

def _gml(t): return f"{{{GML_NS}}}{t}"
def lbl(val,lookup):
    if pd.isna(val): return ""
    try: return lookup.get(int(val),val)
    except: return val
def bool_to_types(row,type_cols):
    s=[TYPE_LABELS.get(c,c) for c in type_cols if str(row.get(c,False)).strip().lower() in ("true","1","yes")]
    return "; ".join(s) if s else ""
def has_attrs(row):
    for col in ["Depth","Frequency","Reciprocity","Trust","Energy","Support","Creativity","Knowledge"]:
        v=row.get(col)
        if v is not None:
            try:
                if not pd.isna(v): return True
            except: return True
    return False
def assign_question(df):
    qs,bare=[],0
    for _,r in df.iterrows():
        if has_attrs(r): qs.append("Current Connections")
        else:
            bare+=1
            qs.append("Aspired Connections" if bare<=MAX_BARE else ("Informal Influence" if bare<=MAX_BARE*2 else "Unknown"))
    df=df.copy(); df["network_question"]=qs; return df

# ── Parsers ───────────────────────────────────────────────────────────────────
def parse_ego(b):
    df=pd.read_csv(io.BytesIO(b),encoding="utf-8-sig",encoding_errors="replace")
    if df.empty: return {}
    r=df.iloc[0]
    return {"ego_uuid":r.get("networkCanvasEgoUUID",""),"case_id":r.get("networkCanvasCaseID",""),
            "name":r.get("Name",""),"geography":r.get("Geography",""),
            "sector":lbl(r.get("Sector"),SECTOR_LABELS),"sharing":lbl(r.get("Sharing"),SHARING_LABELS),
            "confidentiality":lbl(r.get("Confidentiality"),CONFIDENTIALITY_LABELS),
            "consent":str(r.get("Consent","")),"expansion":r.get("Expansion",""),
            "open":r.get("Open",""),"tenure":r.get("Tenure",""),
            "session_start":r.get("sessionStart",""),"session_finish":r.get("sessionFinish","")}

def parse_attributes(b,ego_uuid,ego_name):
    df=pd.read_csv(io.BytesIO(b),encoding="utf-8-sig",encoding_errors="replace"); df=assign_question(df)
    type_cols=[c for c in df.columns if c.startswith("Type_")]
    nodes=[]
    for _,r in df.iterrows():
        nodes.append({"ego_uuid":ego_uuid,"ego_name":ego_name,"node_uuid":r.get("networkCanvasUUID",""),
            "name":r.get("Name",""),"organization":r.get("Organization",""),
            "network_question":r.get("network_question",""),
            "frequency":lbl(r.get("Frequency"),FREQUENCY_LABELS),"depth":lbl(r.get("Depth"),DEPTH_LABELS),
            "knowledge":lbl(r.get("Knowledge"),KNOWLEDGE_LABELS),"energy":lbl(r.get("Energy"),SCALE4_LABELS),
            "support":lbl(r.get("Support"),SCALE4_LABELS),"creativity":lbl(r.get("Creativity"),SCALE4_LABELS),
            "trust":lbl(r.get("Trust"),SCALE4_LABELS),"relationship_type":bool_to_types(r,type_cols),
            "reciprocity":lbl(r.get("Reciprocity"),RECIPROCITY_LABELS)})
    return nodes

def parse_graphml(b,ego_uuid,ego_name,uuid_to_name):
    tree=ET.parse(io.BytesIO(b)); root=tree.getroot()
    key_map={k.attrib.get("id",""):k.attrib.get("attr.name","") for k in root.findall(_gml("key"))}
    graph=root.find(_gml("graph"))
    if graph is None: return []
    node_data={}
    for node in graph.findall(_gml("node")):
        nid=node.attrib.get("id","")
        props={key_map.get(d.attrib.get("key",""),d.attrib.get("key","")):d.text or "" for d in node.findall(_gml("data"))}
        node_data[nid]={"uuid":props.get("networkCanvasUUID",""),"name":props.get("Name","")}
        if props.get("networkCanvasUUID"): uuid_to_name[props["networkCanvasUUID"]]=props.get("Name","")
    edges=[]
    for edge in graph.findall(_gml("edge")):
        props={key_map.get(d.attrib.get("key",""),d.attrib.get("key","")):d.text or "" for d in edge.findall(_gml("data"))}
        sn=node_data.get(edge.attrib.get("source",""),{}); tn=node_data.get(edge.attrib.get("target",""),{})
        edges.append({"From":sn.get("name",""),"To":tn.get("name","")})
    return edges

def extract_zips(uploaded_zips):
    buckets={}
    for f in uploaded_zips:
        raw=f.read()
        if not zipfile.is_zipfile(io.BytesIO(raw)): st.warning(f"⚠️ `{f.name}` is not a valid zip — skipped."); continue
        with zipfile.ZipFile(io.BytesIO(raw)) as zf:
            for zname in zf.namelist():
                if zname.endswith("/") or "__MACOSX" in zname.split("/"): continue
                basename=zname.split("/")[-1]
                if not basename or basename.startswith("._"): continue
                fb=zf.read(zname)
                if basename.endswith("_ego.csv"): buckets.setdefault(basename.replace("_ego.csv",""),{})["ego"]=fb
                elif basename.endswith("_attributeList_Person.csv"): buckets.setdefault(basename.replace("_attributeList_Person.csv",""),{})["attrs"]=fb
                elif basename.endswith("_edgeList.csv"): buckets.setdefault(basename.replace("_edgeList.csv",""),{})["edges_csv"]=fb
                elif basename.endswith(".graphml"): buckets.setdefault(basename.replace(".graphml",""),{})["graphml"]=fb
    return buckets

def process_buckets(buckets):
    all_egos,all_edges,warnings=[],[],[]
    uuid_to_name={}
    for sid,files in buckets.items():
        if "ego" not in files: warnings.append(f"No ego file for `{sid}` — skipped."); continue
        ego=parse_ego(files["ego"])
        ego_uuid=ego.get("ego_uuid",sid); ego_name=ego.get("name") or ego.get("case_id") or sid
        uuid_to_name[ego_uuid]=ego_name; all_egos.append(ego)
        nodes=[]
        if "attrs" in files: nodes=parse_attributes(files["attrs"],ego_uuid,ego_name)
        for n in nodes:
            if n["network_question"]!="Current Connections": continue
            all_edges.append({"From":ego_name,"To":n["name"],"To Organization":n["organization"],
                "Frequency of Interaction":n["frequency"],"Depth of Connection":n["depth"],
                "Knowledge of Their Work":n["knowledge"],"Energy":n["energy"],"Support":n["support"],
                "Creativity":n["creativity"],"Trust":n["trust"],"Relationship Type(s)":n["relationship_type"],
                "Reciprocity":n["reciprocity"]})
        if "graphml" in files: parse_graphml(files["graphml"],ego_uuid,ego_name,uuid_to_name)
    ego_df=pd.DataFrame(all_egos).drop(columns=["ego_uuid","case_id"],errors="ignore").rename(columns={
        "name":"Name","geography":"Geography","sector":"Sector",
        "sharing":"Dimension Most Excited to Share","confidentiality":"Confidentiality Preference",
        "consent":"Consent","expansion":"Expansion / Other Members","open":"Open to New Members?",
        "tenure":"Tenure","session_start":"Session Start","session_finish":"Session Finish"})
    edge_df=pd.DataFrame(all_edges) if all_edges else pd.DataFrame()
    return ego_df,edge_df,warnings

def parse_presurvey(uploaded_file):
    df=pd.read_excel(uploaded_file,sheet_name="Form Responses 1")
    df=df.rename(columns={
        "Name":"name","Organization Name":"organization","Sector":"sector_raw",
        "Skills you personally can offer (to help others in this network)":"skills_offered",
        "Skills you (or your org) need help with (to accomplish your mission)":"skills_needed",
        "Who is your org's primary target audience? (i.e., who are you serving? The more specific, the better)":"target_audience",
        "How would you describe the most challenging issue you (or your org) face? ":"challenge",
        "How would you describe your context? [Historic / Stagnant]":"ctx_historic",
        "How would you describe your context? [Transitional]":"ctx_transitional",
        "How would you describe your context? [Affluent ]":"ctx_affluent",
        "How would you describe your context? [University / Transient]":"ctx_university",
    })
    def parse_contexts(row):
        found=[]
        for col,lbl_ in [("ctx_historic","Historic/Stagnant"),("ctx_transitional","Transitional"),
                          ("ctx_affluent","Affluent"),("ctx_university","University/Transient")]:
            if pd.notna(row.get(col,"")): found.append(lbl_)
        return ", ".join(found) if found else ""
    df["neighborhood_contexts"]=df.apply(parse_contexts,axis=1)
    df["name_clean"]=df["name"].str.strip().str.lower()
    return df[["name","name_clean","organization","sector_raw","skills_offered","skills_needed","challenge","neighborhood_contexts","target_audience"]]

def merge_presurvey(ego_df,pre_df):
    if ego_df.empty or pre_df.empty: return ego_df
    ego_df=ego_df.copy()
    ego_df["name_clean"]=ego_df["Name"].str.strip().str.lower()
    merged=ego_df.merge(pre_df[["name_clean","skills_offered","skills_needed","challenge","neighborhood_contexts","sector_raw"]],
                        on="name_clean",how="left").drop(columns=["name_clean"])
    return merged

# ── Excel builder ─────────────────────────────────────────────────────────────
def build_excel(ego_df,edge_df):
    buf=io.BytesIO()
    with pd.ExcelWriter(buf,engine="openpyxl") as w:
        ego_df.to_excel(w,sheet_name="Nodes",index=False)
        if not edge_df.empty: edge_df.to_excel(w,sheet_name="Edges",index=False)
    buf.seek(0); wb=load_workbook(buf)
    hf=PatternFill("solid",start_color="2825BE"); hfont=Font(name="Arial",bold=True,color="FFFFFF",size=10)
    bf=Font(name="Arial",size=10); alt=PatternFill("solid",start_color="F5F7F6")
    ca=Alignment(horizontal="center",vertical="top",wrap_text=True)
    la=Alignment(horizontal="left",vertical="top",wrap_text=True)
    tb=Border(bottom=Side(style="thin",color="e2e5e3"),right=Side(style="thin",color="e2e5e3"))
    for ws in wb.worksheets:
        for cell in ws[1]: cell.font=hfont; cell.fill=hf; cell.alignment=ca; cell.border=tb
        ws.freeze_panes="A2"; ws.row_dimensions[1].height=30
        for ri,row in enumerate(ws.iter_rows(min_row=2),start=2):
            fill=alt if ri%2==0 else None
            for cell in row:
                cell.font=bf; cell.alignment=la; cell.border=tb
                if fill: cell.fill=fill
        for cc in ws.columns:
            ml=max((len(str(c.value)) if c.value else 0) for c in cc)
            ws.column_dimensions[get_column_letter(cc[0].column)].width=min(ml+4,50)
    out=io.BytesIO(); wb.save(out); out.seek(0); return out

# ── Graph ─────────────────────────────────────────────────────────────────────
def build_graph(edge_df):
    G=nx.DiGraph()
    if edge_df.empty: return G
    for _,r in edge_df.iterrows():
        frm=str(r["From"]).strip(); to=str(r["To"]).strip()
        if not frm or not to or frm in ("nan","None") or to in ("nan","None"): continue
        if frm==to: continue  # skip self-loops
        if G.has_edge(frm,to): G[frm][to]["weight"]=G[frm][to].get("weight",1)+1
        else: G.add_edge(frm,to,weight=1)
    return G

# ── Network figure (light bg, visible edges) ──────────────────────────────────
def make_network_fig(G,color_by="degree",ego_names=None,node_meta=None,show_labels=True):
    if len(G.nodes)==0: return go.Figure()
    UG=G.to_undirected()
    pos=nx.spring_layout(UG,seed=42,k=2.8/math.sqrt(max(len(G.nodes),1)))
    degree=dict(G.degree()); bw=nx.betweenness_centrality(G); indeg=dict(G.in_degree())
    ego_set=set(ego_names or []); max_deg=max(degree.values()) if degree else 1

    def get_color(n):
        if color_by=="degree":
            v=degree.get(n,0); t=0.3+0.7*(v/max(max_deg,1))
            return f"rgba(40,37,190,{t:.2f})"
        elif color_by=="betweenness":
            v=bw.get(n,0); mv=max(bw.values()) if bw else 1; t=0.3+0.7*(v/max(mv,0.001))
            return f"rgba(12,122,122,{t:.2f})"
        elif color_by=="in_degree":
            v=indeg.get(n,0); mi=max(indeg.values()) if indeg else 1; t=0.3+0.7*(v/max(mi,1))
            return f"rgba(235,144,1,{t:.2f})"
        elif color_by=="survey_taker": return AMBER if n in ego_set else INDIGO
        elif color_by=="isolated": return TERRA if indeg.get(n,0)==0 and n not in ego_set else INDIGO
        return INDIGO

    node_colors=[get_color(n) for n in G.nodes]
    node_sizes=[12+30*(degree.get(n,0)/max(max_deg,1)) for n in G.nodes]

    def hover(n):
        lines=[f"<b>{n}</b>","",f"Connections: {degree.get(n,0)}",f"Times named: {indeg.get(n,0)}",f"Betweenness: {bw.get(n,0):.3f}"]
        if node_meta and n in node_meta:
            m=node_meta[n]
            if m.get("challenge"): lines+=["",f"<i>Challenge:</i> {str(m['challenge'])[:120]}"]
            if m.get("skills_offered"): lines+=[f"<i>Offers:</i> {str(m['skills_offered'])[:100]}"]
        return "<br>".join(lines)

    ex,ey=[],[]
    for u,v in G.edges():
        x0,y0=pos[u]; x1,y1=pos[v]
        ex+=[x0,x1,None]; ey+=[y0,y1,None]

    edge_trace=go.Scatter(x=ex,y=ey,mode="lines",
        line=dict(width=1.2,color="rgba(40,37,190,0.25)"),hoverinfo="none",showlegend=False)
    mode="markers+text" if show_labels else "markers"
    node_trace=go.Scatter(
        x=[pos[n][0] for n in G.nodes],y=[pos[n][1] for n in G.nodes],
        mode=mode,text=[n for n in G.nodes],textposition="top center",
        textfont=dict(family="IBM Plex Mono",size=9,color=TEXT2),
        hovertext=[hover(n) for n in G.nodes],hoverinfo="text",
        marker=dict(size=node_sizes,color=node_colors,line=dict(width=2,color="white")),
        showlegend=False)

    fig=go.Figure(data=[edge_trace,node_trace])
    fig.update_layout(
        paper_bgcolor=SURFACE2,plot_bgcolor=SURFACE2,
        margin=dict(l=16,r=16,t=16,b=16),
        xaxis=dict(showgrid=False,zeroline=False,showticklabels=False,showline=False),
        yaxis=dict(showgrid=False,zeroline=False,showticklabels=False,showline=False),
        height=560,font=dict(family="IBM Plex Sans"),
        hoverlabel=dict(bgcolor=INK2,font_color="white",font_family="IBM Plex Sans",align="left"))
    return fig

# ── Analytics helpers ─────────────────────────────────────────────────────────
def pct_high(series):
    counts=series.value_counts()
    high=counts.get("A great deal",0)+counts.get("Quite a bit",0)
    return round(100*high/max(counts.sum(),1))

def pct_label(val,series,order):
    counts=series.value_counts()
    total=counts.sum()
    return round(100*counts.get(val,0)/max(total,1))

def parse_skills(s):
    if pd.isna(s) or not str(s).strip(): return []
    return [x.strip() for x in str(s).split(",") if x.strip()]

def build_skill_match(pre_df):
    off=collections.defaultdict(list); need=collections.defaultdict(list)
    for _,r in pre_df.iterrows():
        n=r.get("name","")
        for sk in parse_skills(r.get("skills_offered","")): off[sk].append(n)
        for sk in parse_skills(r.get("skills_needed","")): need[sk].append(n)
    all_sk=set(off)|set(need)
    return sorted([{"skill":sk,"offered":len(off.get(sk,[])),"needed":len(need.get(sk,[])),
                    "gap":len(need.get(sk,[]))-len(off.get(sk,[])),
                    "offerers":off.get(sk,[]),"needers":need.get(sk,[])} for sk in all_sk],
                  key=lambda x:-x["needed"])

# ── Network analytics summary (for both app + HTML export) ────────────────────
def compute_analytics(ego_df,edge_df,G,pre_df=None):
    """Returns a structured dict of key findings."""
    a={}
    if edge_df.empty: return a
    ego_set=set(ego_df["Name"].dropna().str.strip()) if "Name" in ego_df.columns else set()
    indeg=dict(G.in_degree()); bw=nx.betweenness_centrality(G)

    a["n_respondents"]=len(ego_df)
    a["n_nodes"]=len(G.nodes)
    a["n_edges"]=len(G.edges)
    a["density"]=round(nx.density(G),3)

    # Depth
    if "Depth of Connection" in edge_df.columns:
        dc=edge_df["Depth of Connection"].value_counts(normalize=True)*100
        a["pct_deep"]=round(dc.get("Collaboration",0)+dc.get("Cooperation",0))
        a["pct_shallow"]=round(dc.get("Awareness",0)+dc.get("Connection",0))
        depth_counts=edge_df["Depth of Connection"].value_counts()
        a["depth_counts"]={k:int(v) for k,v in depth_counts.items()}

    # Relational quality
    for attr in ["Energy","Support","Creativity","Trust","Knowledge"]:
        if attr in edge_df.columns:
            a[f"pct_{attr.lower()}"]=pct_high(edge_df[attr])

    # Low-trust edges
    if "Trust" in edge_df.columns:
        low_trust=(edge_df["Trust"].isin(["Not at all","Somewhat"])).sum()
        a["n_low_trust"]=int(low_trust)
        a["pct_low_trust"]=round(100*low_trust/max(len(edge_df),1))

    # Low-energy edges
    if "Energy" in edge_df.columns:
        low_e=(edge_df["Energy"].isin(["Not at all","Somewhat"])).sum()
        a["n_low_energy"]=int(low_e)
        a["pct_low_energy"]=round(100*low_e/max(len(edge_df),1))

    # Reciprocity
    if "Reciprocity" in edge_df.columns:
        a["pct_balanced"]=round(100*(edge_df["Reciprocity"]=="Roughly balanced").sum()/max(len(edge_df),1))
        a["pct_one_way"]=round(100*(edge_df["Reciprocity"].isin(["Mostly flows from them to me","Mostly flows from me to them"])).sum()/max(len(edge_df),1))

    # Frequency
    if "Frequency of Interaction" in edge_df.columns:
        fc=edge_df["Frequency of Interaction"].value_counts(normalize=True)*100
        a["pct_frequent"]=round(fc.get("Weekly",0)+fc.get("Multiple times a week",0)+fc.get("Daily or near-daily",0))
        a["pct_infrequent"]=round(fc.get("Never or rarely",0)+fc.get("Once or twice a month",0))

    # Relationship types
    if "Relationship Type(s)" in edge_df.columns:
        all_t=[]
        for v in edge_df["Relationship Type(s)"].dropna():
            all_t.extend([t.strip() for t in str(v).split(";") if t.strip()])
        if all_t:
            tc=collections.Counter(all_t)
            a["top_rel_type"]=tc.most_common(1)[0][0]
            a["rel_type_counts"]=dict(tc.most_common(6))

    # Isolation
    survey_lonely=[n for n in ego_set if indeg.get(n,0)==0]
    periphery=[n for n in G.nodes if indeg.get(n,0)==0 and n not in ego_set]
    a["n_survey_isolated"]=len(survey_lonely)
    a["n_periphery"]=len(periphery)

    # Bridges
    top_bw=sorted(bw.items(),key=lambda x:-x[1])
    if top_bw: a["top_bridge"]=top_bw[0][0]; a["top_bridge_score"]=round(top_bw[0][1],3)

    # Most named
    if "To" in edge_df.columns:
        top=edge_df["To"].value_counts().head(5)
        a["most_named"]=[(n,int(c)) for n,c in top.items()]

    # Geographic spread
    if "Geography" in ego_df.columns:
        geos=ego_df["Geography"].dropna()
        a["n_geographies"]=int(geos.nunique())
        a["geo_counts"]=dict(geos.value_counts())

    # Communities
    try:
        UG_clean=nx.Graph()
        for u,v in G.to_undirected().edges():
            if isinstance(u,str) and isinstance(v,str) and u!=v:
                UG_clean.add_edge(u,v)
        for n in G.nodes():
            if isinstance(n,str): UG_clean.add_node(n)
        comms=nx.community.louvain_communities(UG_clean,seed=42)
        a["n_communities"]=len(comms)
        a["communities"]=[sorted(list(c)) for c in sorted(comms,key=len,reverse=True)]
    except: a["n_communities"]=1; a["communities"]=[]

    # Skills
    if pre_df is not None and not pre_df.empty:
        matches=build_skill_match(pre_df)
        a["skill_matches"]=matches
        gaps=[m for m in matches if m["gap"]>0]
        a["top_skill_gaps"]=gaps[:3]
        surplus=[m for m in matches if m["gap"]<0]
        a["top_skill_surplus"]=surplus[:2]

    return a

# ── HTML insights generator ───────────────────────────────────────────────────
def build_insights_html(a,ego_df,edge_df):
    """Build a projectable, anonymized HTML insights document."""
    n_resp=a.get("n_respondents",0)
    n_nodes=a.get("n_nodes",0)
    n_edges=a.get("n_edges",0)
    density=a.get("density",0)
    pct_deep=a.get("pct_deep",0)
    pct_trust=a.get("pct_trust",0)
    pct_energy=a.get("pct_energy",0)
    pct_creativity=a.get("pct_creativity",0)
    pct_support=a.get("pct_support",0)
    pct_low_trust=a.get("pct_low_trust",0)
    pct_low_energy=a.get("pct_low_energy",0)
    pct_balanced=a.get("pct_balanced",0)
    pct_frequent=a.get("pct_frequent",0)
    n_communities=a.get("n_communities",0)
    n_geographies=a.get("n_geographies",0)
    n_isolated=a.get("n_survey_isolated",0)
    n_periphery=a.get("n_periphery",0)
    top_rel=a.get("top_rel_type","Peer Learning / Knowledge Exchange")

    # Sectors
    sector_html=""
    if "Sector" in ego_df.columns:
        sc=ego_df["Sector"].value_counts()
        sector_html="".join([f'<div class="chip">{s} <span class="chip-n">{c}</span></div>' for s,c in sc.items()])

    # Geo
    geo_html=""
    geo_counts=a.get("geo_counts",{})
    if geo_counts:
        geo_html="".join([f'<div class="geo-row"><span class="geo-name">{g}</span><span class="geo-bar"><span class="geo-fill" style="width:{min(100,int(c/max(geo_counts.values())*100))}%"></span></span><span class="geo-n">{c}</span></div>' for g,c in list(geo_counts.items())[:8]])

    # Skill gaps
    skill_html=""
    for m in a.get("top_skill_gaps",[]):
        need=m["needed"]; off=m["offered"]; gap=m["gap"]
        skill_html+=f'<div class="skill-row"><div class="skill-name">{m["skill"]}</div><div class="skill-bars"><div class="skill-bar-wrap"><div class="skill-bar-fill" style="width:{min(100,need*16)}%;background:#EB9001;"></div></div><div class="skill-bar-wrap"><div class="skill-bar-fill" style="width:{min(100,off*16)}%;background:#0C7A7A;"></div></div></div><div class="skill-meta">{need} need · {off} offer</div></div>'

    # Communities
    comm_html=""
    for i,c in enumerate(a.get("communities",[])[:4]):
        size=len(c)
        comm_html+=f'<div class="comm-card"><div class="comm-num">Cluster {i+1}</div><div class="comm-size">{size} member{"s" if size!=1 else ""}</div></div>'

    # Questions for the room
    questions=[]
    if pct_low_trust>20:
        questions.append(("Trust", f"{pct_low_trust}% of connections score low on trust.", "What would it take to deepen trust in relationships that feel thin? What has worked in your context?"))
    if pct_low_energy>25:
        questions.append(("Energy", f"{pct_low_energy}% of connections feel low-energy.", "Which relationships feel draining rather than generative? What conditions tend to produce energy in your work?"))
    if pct_balanced<50:
        questions.append(("Reciprocity", f"Many connections flow primarily in one direction.", "Where do you notice imbalance in your working relationships? What would more mutual exchange look like?"))
    if n_isolated>0:
        questions.append(("Isolation", f"Some members of this network have not yet been named by others.", "Who in this room do you know least well? What has kept you from connecting?"))
    if n_geographies>2:
        questions.append(("Geography", f"This network spans {n_geographies} geographic areas.", "Where does distance make collaboration harder? What practices help you stay connected across geography?"))
    if n_communities>1:
        questions.append(("Clusters", f"The network appears to have {n_communities} distinct clusters.", "Which groups in this room interact least with each other? What sits between them?"))
    if pct_deep<50:
        questions.append(("Depth", f"Most connections are still at early stages of depth.", "Which relationships here have real potential to deepen? What would cooperation or collaboration actually look like?"))

    if not questions:
        questions=[("Ecosystem","The network is taking shape.",
                    "Who in this room do you know least well? What would it mean to strengthen that connection?")]

    q_html="".join([f'''<div class="question-card">
        <div class="q-tag" style="background:{"#CF4C38" if qt in ["Trust","Isolation","Energy"] else "#EB9001"}">{qt}</div>
        <div class="q-stat">{qs}</div>
        <div class="q-prompt">{qp}</div>
    </div>''' for qt,qs,qp in questions])

    html=f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>TNN Network Insights</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@700;800;900&family=IBM+Plex+Sans:wght@300;400;500;600&family=IBM+Plex+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
:root{{--indigo:#2825BE;--amber:#EB9001;--teal:#0C7A7A;--terra:#CF4C38;--green:#2E9E5B;
  --ink:#080818;--ink2:#0f0e2e;--ink3:#13123a;--s2:#F5F7F6;--border:#e2e5e3;
  --text:#111827;--text2:#4b5563;--text3:#9ca3af;}}
*{{box-sizing:border-box;margin:0;padding:0;}}
body{{font-family:'IBM Plex Sans',sans-serif;background:var(--ink);color:white;}}
.page{{max-width:1120px;margin:0 auto;padding:64px 48px;}}

/* COVER */
.cover{{min-height:100vh;display:flex;flex-direction:column;justify-content:flex-end;padding:96px 48px;background:var(--ink);border-bottom:1px solid rgba(255,255,255,0.08);}}
.cover-eyebrow{{font-family:'IBM Plex Mono',monospace;font-size:11px;letter-spacing:0.18em;text-transform:uppercase;color:var(--amber);margin-bottom:24px;}}
.cover-title{{font-family:'Barlow Condensed',sans-serif;font-size:clamp(52px,7vw,96px);font-weight:900;line-height:0.95;letter-spacing:-0.02em;text-transform:uppercase;color:white;margin-bottom:24px;}}
.cover-sub{{font-family:'IBM Plex Sans',sans-serif;font-size:18px;color:rgba(255,255,255,0.5);max-width:520px;line-height:1.6;}}
.cover-meta{{font-family:'IBM Plex Mono',monospace;font-size:11px;color:rgba(255,255,255,0.3);margin-top:48px;letter-spacing:0.1em;}}

/* SECTION */
.section{{padding:80px 0;border-bottom:1px solid rgba(255,255,255,0.08);}}
.section:last-child{{border-bottom:none;}}
.section-eyebrow{{font-family:'IBM Plex Mono',monospace;font-size:10px;letter-spacing:0.18em;text-transform:uppercase;color:var(--amber);margin-bottom:12px;}}
.section-title{{font-family:'Barlow Condensed',sans-serif;font-size:clamp(32px,4vw,52px);font-weight:800;text-transform:uppercase;letter-spacing:-0.01em;color:white;margin-bottom:32px;line-height:1.1;}}

/* STAT GRID */
.stat-grid{{display:grid;grid-template-columns:repeat(4,1fr);gap:2px;margin:32px 0;}}
.stat-cell{{background:var(--ink2);padding:32px 24px;}}
.stat-cell .sv{{font-family:'Barlow Condensed',sans-serif;font-size:56px;font-weight:900;line-height:1;color:var(--indigo);}}
.stat-cell .sl{{font-family:'IBM Plex Mono',monospace;font-size:10px;color:var(--text3);text-transform:uppercase;letter-spacing:0.1em;margin-top:6px;}}

/* QUALITY SCORES */
.quality-grid{{display:grid;grid-template-columns:repeat(2,1fr);gap:2px;margin:24px 0;}}
.quality-cell{{background:var(--ink2);padding:24px;display:flex;align-items:center;gap:20px;}}
.quality-cell .qv{{font-family:'Barlow Condensed',sans-serif;font-size:44px;font-weight:800;min-width:80px;line-height:1;}}
.quality-cell .qinfo .ql{{font-family:'IBM Plex Mono',monospace;font-size:10px;color:var(--text3);text-transform:uppercase;letter-spacing:0.1em;}}
.quality-cell .qinfo .qsub{{font-family:'IBM Plex Sans',sans-serif;font-size:13px;color:rgba(255,255,255,0.6);margin-top:4px;line-height:1.5;}}

/* DEPTH BAR */
.depth-bar{{background:var(--ink2);padding:32px;margin:16px 0;border-radius:2px;}}
.depth-bar-label{{font-family:'IBM Plex Mono',monospace;font-size:10px;color:var(--text3);text-transform:uppercase;letter-spacing:0.1em;margin-bottom:16px;}}
.depth-segments{{display:flex;height:40px;border-radius:2px;overflow:hidden;gap:2px;}}
.depth-seg{{display:flex;align-items:center;justify-content:center;font-family:'IBM Plex Mono',monospace;font-size:10px;color:white;transition:flex 0.3s;}}
.depth-legend{{display:flex;gap:24px;margin-top:12px;}}
.dl-item{{display:flex;align-items:center;gap:6px;font-family:'IBM Plex Sans',sans-serif;font-size:12px;color:rgba(255,255,255,0.5);}}
.dl-dot{{width:8px;height:8px;border-radius:50%;}}

/* CHIPS */
.chip-row{{display:flex;flex-wrap:wrap;gap:8px;margin:16px 0;}}
.chip{{background:var(--ink2);border:1px solid rgba(255,255,255,0.1);padding:6px 12px;border-radius:2px;font-family:'IBM Plex Sans',sans-serif;font-size:13px;color:rgba(255,255,255,0.7);display:flex;align-items:center;gap:8px;}}
.chip-n{{font-family:'IBM Plex Mono',monospace;font-size:11px;color:var(--amber);}}

/* GEO */
.geo-row{{display:flex;align-items:center;gap:12px;margin:8px 0;}}
.geo-name{{font-family:'IBM Plex Sans',sans-serif;font-size:13px;color:rgba(255,255,255,0.7);min-width:160px;}}
.geo-bar{{flex:1;height:6px;background:rgba(255,255,255,0.08);border-radius:3px;overflow:hidden;}}
.geo-fill{{height:100%;background:var(--indigo);border-radius:3px;}}
.geo-n{{font-family:'IBM Plex Mono',monospace;font-size:11px;color:var(--text3);min-width:24px;text-align:right;}}

/* SKILLS */
.skill-row{{margin:16px 0;}}
.skill-name{{font-family:'IBM Plex Sans',sans-serif;font-size:14px;color:white;margin-bottom:6px;}}
.skill-bars{{display:flex;flex-direction:column;gap:3px;margin-bottom:4px;}}
.skill-bar-wrap{{height:8px;background:rgba(255,255,255,0.06);border-radius:2px;overflow:hidden;width:100%;}}
.skill-bar-fill{{height:100%;border-radius:2px;}}
.skill-meta{{font-family:'IBM Plex Mono',monospace;font-size:10px;color:var(--text3);}}

/* COMMUNITIES */
.comm-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:2px;margin:24px 0;}}
.comm-card{{background:var(--ink2);padding:24px;}}
.comm-num{{font-family:'IBM Plex Mono',monospace;font-size:10px;color:var(--amber);text-transform:uppercase;letter-spacing:0.1em;margin-bottom:8px;}}
.comm-size{{font-family:'Barlow Condensed',sans-serif;font-size:32px;font-weight:800;color:white;}}

/* QUESTION CARDS */
.question-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:2px;margin:24px 0;}}
.question-card{{background:var(--ink2);padding:32px;}}
.q-tag{{display:inline-block;font-family:'IBM Plex Mono',monospace;font-size:9px;letter-spacing:0.14em;text-transform:uppercase;color:white;padding:3px 8px;border-radius:2px;margin-bottom:16px;}}
.q-stat{{font-family:'IBM Plex Sans',sans-serif;font-size:13px;color:rgba(255,255,255,0.5);margin-bottom:12px;line-height:1.5;}}
.q-prompt{{font-family:'IBM Plex Sans',sans-serif;font-size:16px;color:white;line-height:1.65;font-style:italic;}}

/* CLOSING */
.closing{{background:var(--ink2);padding:64px 48px;margin-top:80px;border-top:3px solid var(--amber);}}
.closing-title{{font-family:'Barlow Condensed',sans-serif;font-size:44px;font-weight:800;text-transform:uppercase;color:white;margin-bottom:16px;line-height:1.1;}}
.closing-body{{font-family:'IBM Plex Sans',sans-serif;font-size:16px;color:rgba(255,255,255,0.6);max-width:600px;line-height:1.7;}}
.footer{{font-family:'IBM Plex Mono',monospace;font-size:10px;color:rgba(255,255,255,0.2);text-align:center;padding:32px;letter-spacing:0.1em;}}
@media print{{body{{-webkit-print-color-adjust:exact;print-color-adjust:exact;}}}}
</style>
</head>
<body>

<div class="cover">
  <div class="cover-eyebrow">Connecting for Change · Thriving Neighborhoods Network</div>
  <div class="cover-title">What the<br>Network<br>Shows</div>
  <div class="cover-sub">An anonymized portrait of connections, strengths, and opportunities across the TNN ecosystem — prepared to support Activity 5: Who Is in the Ecosystem?</div>
  <div class="cover-meta">n={n_resp} respondents · {n_nodes} nodes · {n_edges} connections · density={density}</div>
</div>

<div class="page">

<!-- SCALE -->
<div class="section">
  <div class="section-eyebrow">Network Scale</div>
  <div class="section-title">The shape<br>of what exists</div>
  <div class="stat-grid">
    <div class="stat-cell"><div class="sv">{n_resp}</div><div class="sl">Respondents</div></div>
    <div class="stat-cell"><div class="sv">{n_nodes}</div><div class="sl">People in the network</div></div>
    <div class="stat-cell"><div class="sv">{n_edges}</div><div class="sl">Named connections</div></div>
    <div class="stat-cell"><div class="sv">{n_communities}</div><div class="sl">Distinct clusters</div></div>
  </div>
  <p style="font-family:'IBM Plex Sans',sans-serif;font-size:15px;color:rgba(255,255,255,0.5);max-width:600px;line-height:1.7;">
    Each respondent named people they work with and rated the quality of those relationships.
    The network you see here is built from those reports — aggregated, anonymized, and combined
    into a single view of the ecosystem.
  </p>
</div>

<!-- DEPTH -->
<div class="section">
  <div class="section-eyebrow">Depth of Connection</div>
  <div class="section-title">{pct_deep}% of connections<br>reach cooperation<br>or collaboration</div>
  <div class="depth-bar">
    <div class="depth-bar-label">Distribution across all reported connections</div>
    <div class="depth-segments">
      {f'<div class="depth-seg" style="flex:{a.get("depth_counts",{{}}).get("Awareness",0)};background:#13123a;">{a.get("depth_counts",{{}}).get("Awareness",0)}</div>' if a.get("depth_counts",{}).get("Awareness",0) else ""}
      {f'<div class="depth-seg" style="flex:{a.get("depth_counts",{{}}).get("Connection",0)};background:#1e1c4a;">{a.get("depth_counts",{{}}).get("Connection",0)}</div>' if a.get("depth_counts",{}).get("Connection",0) else ""}
      {f'<div class="depth-seg" style="flex:{a.get("depth_counts",{{}}).get("Cooperation",0)};background:#2825BE;">{a.get("depth_counts",{{}}).get("Cooperation",0)}</div>' if a.get("depth_counts",{}).get("Cooperation",0) else ""}
      {f'<div class="depth-seg" style="flex:{a.get("depth_counts",{{}}).get("Collaboration",0)};background:#4a47d6;">{a.get("depth_counts",{{}}).get("Collaboration",0)}</div>' if a.get("depth_counts",{}).get("Collaboration",0) else ""}
    </div>
    <div class="depth-legend">
      <div class="dl-item"><div class="dl-dot" style="background:#13123a;"></div>Awareness</div>
      <div class="dl-item"><div class="dl-dot" style="background:#1e1c4a;"></div>Connection</div>
      <div class="dl-item"><div class="dl-dot" style="background:#2825BE;"></div>Cooperation</div>
      <div class="dl-item"><div class="dl-dot" style="background:#4a47d6;"></div>Collaboration</div>
    </div>
  </div>
</div>

<!-- RELATIONAL QUALITY -->
<div class="section">
  <div class="section-eyebrow">Relational Quality</div>
  <div class="section-title">% of connections rated<br>"quite a bit" or "a great deal"</div>
  <div class="quality-grid">
    <div class="quality-cell"><div class="qv" style="color:{'#2E9E5B' if pct_trust>=60 else '#CF4C38'}">{pct_trust}%</div><div class="qinfo"><div class="ql">Trust</div><div class="qsub">{'Most connections feel trustworthy.' if pct_trust>=60 else f'{pct_low_trust}% of connections score low on trust — a signal worth exploring.'}</div></div></div>
    <div class="quality-cell"><div class="qv" style="color:{'#2E9E5B' if pct_energy>=60 else '#CF4C38'}">{pct_energy}%</div><div class="qinfo"><div class="ql">Energy</div><div class="qsub">{'Connections feel generative and activating.' if pct_energy>=60 else f'{pct_low_energy}% of connections feel low-energy — which relationships could be better resourced?'}</div></div></div>
    <div class="quality-cell"><div class="qv" style="color:{'#2E9E5B' if pct_support>=60 else '#EB9001'}">{pct_support}%</div><div class="qinfo"><div class="ql">Support</div><div class="qsub">{'Strong sense of mutual support across the network.' if pct_support>=60 else 'Support is unevenly distributed — some relationships carry more weight.'}</div></div></div>
    <div class="quality-cell"><div class="qv" style="color:{'#2E9E5B' if pct_creativity>=60 else '#EB9001'}">{pct_creativity}%</div><div class="qinfo"><div class="ql">Creativity</div><div class="qsub">{'Connections feel creatively alive.' if pct_creativity>=60 else 'Creative spark is lower than other dimensions — worth asking what makes some relationships more generative.'}</div></div></div>
  </div>
  <p style="font-family:'IBM Plex Mono',monospace;font-size:11px;color:rgba(255,255,255,0.3);margin-top:12px;">
    Reciprocity: {pct_balanced}% of connections described as roughly balanced · {pct_frequent}% interact weekly or more
  </p>
</div>

<!-- WHO IS IN THE ECOSYSTEM -->
<div class="section">
  <div class="section-eyebrow">Ecosystem Composition · Activity 5</div>
  <div class="section-title">Who is visible<br>in this network</div>
  {'<div class="section-eyebrow" style="color:#0C7A7A;margin-top:24px;">Sectors represented</div><div class="chip-row">'+sector_html+'</div>' if sector_html else ''}
  {'<div class="section-eyebrow" style="color:#0C7A7A;margin-top:32px;">Geographic spread · '+str(n_geographies)+' area'+('s' if n_geographies!=1 else '')+'</div>'+geo_html if geo_html else ''}
  {'<div class="section-eyebrow" style="color:#0C7A7A;margin-top:32px;">Network clusters</div><div class="comm-grid">'+comm_html+'</div>' if comm_html else ''}
  <p style="font-family:'IBM Plex Sans',sans-serif;font-size:14px;color:rgba(255,255,255,0.4);margin-top:24px;max-width:600px;line-height:1.7;">
    The dominant relationship type is <strong style="color:rgba(255,255,255,0.7);">{top_rel}</strong>.
    {"The network spans " + str(n_geographies) + " geographic areas — distance shapes who stays connected." if n_geographies>2 else ""}
  </p>
</div>

<!-- GAPS -->
<div class="section">
  <div class="section-eyebrow" style="color:#CF4C38;">Gaps & Opportunities</div>
  <div class="section-title">Where the network<br>could be stronger</div>
  {'<div style="margin:24px 0;"><div class="section-eyebrow" style="color:#EB9001;margin-bottom:12px;">Skill gaps — what people need that others could offer</div>'+skill_html+'</div>' if skill_html else ''}
  {'<div style="background:#CF4C38;display:inline-block;padding:4px 12px;font-family:IBM Plex Mono,monospace;font-size:11px;color:white;letter-spacing:0.1em;text-transform:uppercase;margin-bottom:12px;">Potential isolation</div><p style="font-family:IBM Plex Sans,sans-serif;font-size:14px;color:rgba(255,255,255,0.6);max-width:580px;line-height:1.7;">'+str(n_isolated+n_periphery)+' people in this network appear in only one person\'s named connections. They may be underconnected relative to their potential role in the ecosystem.</p>' if (n_isolated+n_periphery)>0 else ''}
</div>

<!-- QUESTIONS FOR THE ROOM -->
<div class="section">
  <div class="section-eyebrow">For the room</div>
  <div class="section-title">What the data<br>can't answer alone</div>
  <p style="font-family:'IBM Plex Sans',sans-serif;font-size:15px;color:rgba(255,255,255,0.5);max-width:580px;line-height:1.7;margin-bottom:32px;">
    These questions emerge from the network data. They are not conclusions — they are starting points.
    The ecosystem becomes visible when people in the room name what the data cannot.
  </p>
  <div class="question-grid">{q_html}</div>
</div>

</div>

<div class="closing">
  <div class="closing-title">The map is incomplete<br>by design.</div>
  <div class="closing-body">
    What you see here represents what people chose to name. The relationships that matter most,
    the actors who are missing, the conditions that are fragile — those only become visible
    when people in this room say what the data cannot.<br><br>
    Your role is not to validate the map. Your role is to improve it.
  </div>
</div>

<div class="footer">TNN NETWORK INSIGHTS · CONNECTING FOR CHANGE · ANONYMIZED · NOT FOR DISTRIBUTION</div>
</body>
</html>"""
    return html

# ── Skill gap chart (single clean one) ───────────────────────────────────────
def skill_gap_chart(matches):
    skills=[m["skill"] for m in matches[:10]]
    offered=[m["offered"] for m in matches[:10]]
    needed=[m["needed"] for m in matches[:10]]
    fig=go.Figure()
    fig.add_trace(go.Bar(name="Needed",y=skills,x=needed,orientation="h",marker_color=AMBER,marker_line_width=0))
    fig.add_trace(go.Bar(name="Offered",y=skills,x=offered,orientation="h",marker_color=TEAL,marker_line_width=0,opacity=0.8))
    fig.update_layout(barmode="overlay",paper_bgcolor=SURFACE,plot_bgcolor=SURFACE,
        title=dict(text="Skills: Supply vs. Demand",font=dict(family="Barlow Condensed",size=18,color=INK)),
        xaxis=dict(tickfont=dict(family="IBM Plex Mono",size=10),gridcolor=BORDER,linecolor=BORDER),
        yaxis=dict(tickfont=dict(family="IBM Plex Sans",size=11),showgrid=False,autorange="reversed"),
        legend=dict(font=dict(family="IBM Plex Sans",size=11),orientation="h",y=1.08),
        margin=dict(l=16,r=16,t=48,b=16),height=max(300,32*len(skills)+80))
    return fig

# ─────────────────────────────────────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────────────────────────────────────
st.markdown('<div class="eyebrow">Connecting for Change</div>',unsafe_allow_html=True)
st.markdown('<h1 style="font-size:38px;margin:0 0 4px 0;font-family:Barlow Condensed,sans-serif;">TNN Network Analyzer</h1>',unsafe_allow_html=True)
st.markdown(f'<p style="color:{TEXT2};font-size:14px;margin:0 0 20px 0;">Upload participant exports · Explore the network · Export insights</p>',unsafe_allow_html=True)
st.markdown(f'<hr style="border:none;border-top:1px solid {BORDER};margin:0 0 24px 0">',unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('<div class="eyebrow">Data Input</div>',unsafe_allow_html=True)
    st.markdown("**NetCanvas exports**")
    uploaded=st.file_uploader("Zip files",accept_multiple_files=True,type=["zip"],label_visibility="collapsed")
    st.markdown("---")
    st.markdown("**Pre-survey** *(optional)*")
    st.markdown(f'<p style="font-size:12px;color:{TEXT2};">Network Activation Survey Excel — adds skills, challenges, and context.</p>',unsafe_allow_html=True)
    presurvey_file=st.file_uploader("Excel",type=["xlsx"],label_visibility="collapsed",key="presurvey")

    if uploaded:
        if st.button("▶  Process",type="primary",use_container_width=True):
            with st.spinner("Processing..."):
                buckets=extract_zips(uploaded)
                ego_df,edge_df,warnings=process_buckets(buckets)
                pre_df=parse_presurvey(presurvey_file) if presurvey_file else pd.DataFrame()
                if not pre_df.empty: ego_df=merge_presurvey(ego_df,pre_df)
                G=build_graph(edge_df)
                ego_names=list(ego_df["Name"].dropna()) if "Name" in ego_df.columns else []
                node_meta={}
                if not pre_df.empty:
                    for _,r in pre_df.iterrows():
                        node_meta[r["name"]]={"challenge":r.get("challenge",""),"skills_offered":r.get("skills_offered","")}
                analytics=compute_analytics(ego_df,edge_df,G,pre_df if not pre_df.empty else None)
                st.session_state.update({"ego_df":ego_df,"edge_df":edge_df,"G":G,
                    "ego_names":ego_names,"pre_df":pre_df,"node_meta":node_meta,"analytics":analytics})
            for w in warnings: st.warning(w)
            st.success(f"✅ {len(ego_df)} participant(s) · {len(edge_df)} edge(s)")

    if "edge_df" in st.session_state and not st.session_state["edge_df"].empty:
        st.markdown("---")
        st.markdown('<div class="eyebrow">Exports</div>',unsafe_allow_html=True)
        excel=build_excel(st.session_state["ego_df"],st.session_state["edge_df"])
        st.download_button("⬇  Excel",data=excel,file_name="tnn_combined_network.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
        if "analytics" in st.session_state:
            html_bytes=build_insights_html(st.session_state["analytics"],
                st.session_state["ego_df"],st.session_state["edge_df"]).encode("utf-8")
            st.download_button("⬇  Insights HTML",data=html_bytes,file_name="tnn_network_insights.html",
                mime="text/html",use_container_width=True)

# ── Gate ──────────────────────────────────────────────────────────────────────
if "ego_df" not in st.session_state:
    st.markdown(f"""<div style="background:{SURFACE2};border-radius:8px;padding:48px 32px;text-align:center;border:1px dashed {BORDER};">
        <div style="font-family:'Barlow Condensed',sans-serif;font-size:22px;color:{TEXT2};font-weight:700;text-transform:uppercase;">Upload zip files in the sidebar to get started</div>
        <div style="font-family:'IBM Plex Sans',sans-serif;font-size:13px;color:{TEXT3};margin-top:8px;">Pre-survey Excel is optional — adds skills, challenges, and neighborhood context</div>
    </div>""",unsafe_allow_html=True)
    st.stop()

ego_df=st.session_state["ego_df"]; edge_df=st.session_state["edge_df"]
G=st.session_state["G"]; ego_names=st.session_state.get("ego_names",[])
pre_df=st.session_state.get("pre_df",pd.DataFrame())
node_meta=st.session_state.get("node_meta",{})
a=st.session_state.get("analytics",{})

# ── Metric strip ──────────────────────────────────────────────────────────────
c1,c2,c3,c4=st.columns(4)
for col,(val,lbl_,color) in zip([c1,c2,c3,c4],[
    (len(ego_df),"SURVEY TAKERS",INDIGO),(len(G.nodes),"UNIQUE NODES",TEAL),
    (len(G.edges),"CONNECTIONS",AMBER),(f"{nx.density(G):.3f}","NETWORK DENSITY",GREEN)]):
    with col:
        st.markdown(f'<div class="metric-card" style="border-left-color:{color};"><div class="value" style="color:{color};">{val}</div><div class="label">{lbl_}</div></div>',unsafe_allow_html=True)

st.markdown("<br>",unsafe_allow_html=True)

# ── Tabs ──────────────────────────────────────────────────────────────────────
tab1,tab2,tab3,tab4,tab5=st.tabs(["🕸  Network","📊  Connections","🛠  Skills","💡  Insights","🖥  Room View"])

# ═══ TAB 1: NETWORK ══════════════════════════════════════════════════════════
with tab1:
    ctrl_l,ctrl_r=st.columns([4,1])
    with ctrl_r:
        st.markdown("<br>",unsafe_allow_html=True)
        color_by=st.selectbox("Color nodes by",["degree","in_degree","betweenness","survey_taker","isolated"],
            format_func=lambda x:{"degree":"Connections (degree)","in_degree":"Times named","betweenness":"Bridge role","survey_taker":"Respondent vs. named","isolated":"Potential isolation"}.get(x,x))
        show_labels=st.checkbox("Show labels",value=True)
    with ctrl_l:
        st.plotly_chart(make_network_fig(G,color_by=color_by,ego_names=ego_names,node_meta=node_meta,show_labels=show_labels),use_container_width=True)

    col_a,col_b=st.columns(2)
    with col_a:
        st.markdown(f'<div class="eyebrow" style="color:{TEAL};">Centrality</div>',unsafe_allow_html=True)
        deg=dict(G.degree()); indeg=dict(G.in_degree()); bw=nx.betweenness_centrality(G)
        cent=pd.DataFrame([{"Name":n,"Degree":deg[n],"Times Named":indeg[n],"Betweenness":round(bw[n],3)} for n in G.nodes])
        st.dataframe(cent.sort_values("Degree",ascending=False).reset_index(drop=True),use_container_width=True,hide_index=True)
    with col_b:
        st.markdown(f'<div class="eyebrow" style="color:{TEAL};">Clusters · Louvain</div>',unsafe_allow_html=True)
        UG2=G.to_undirected()
        if len(UG2.nodes)>1:
            try:
                UG2_clean=nx.Graph()
                for u,v in UG2.edges():
                    if isinstance(u,str) and isinstance(v,str) and u!=v:
                        UG2_clean.add_edge(u,v)
                for n in UG2.nodes():
                    if isinstance(n,str): UG2_clean.add_node(n)
                comms=nx.community.louvain_communities(UG2_clean,seed=42)
                rows=[{"Cluster":f"Cluster {i+1}","Size":len(c),"Members":", ".join(sorted(list(c)))} for i,c in enumerate(sorted(comms,key=len,reverse=True))]
                st.dataframe(pd.DataFrame(rows),use_container_width=True,hide_index=True)
            except: st.info("Community detection unavailable.")

# ═══ TAB 2: CONNECTIONS ══════════════════════════════════════════════════════
with tab2:
    if edge_df.empty: st.info("No edge data.")
    else:
        # Scorecards for relational quality
        attrs=["Trust","Energy","Support","Creativity"]
        scores=[a.get(f"pct_{x.lower()}",0) for x in attrs]
        colors=[GREEN if s>=60 else (AMBER if s>=40 else TERRA) for s in scores]
        sub_labels=[]
        for attr,s in zip(attrs,scores):
            if attr=="Trust": sub_labels.append("high-trust" if s>=60 else f"{a.get('pct_low_trust',0)}% score low")
            elif attr=="Energy": sub_labels.append("feel generative" if s>=60 else f"{a.get('pct_low_energy',0)}% feel low-energy")
            else: sub_labels.append("" )

        sc_html='<div class="scorecard-row">'
        sc_html = '<div class="scorecard-row">'
        for attr,s,c,sub in zip(attrs,scores,colors,sub_labels):
            sub_div = f'<div class="sc-sub">{sub}</div>' if sub else ''
            sc_html += f'<div class="scorecard" style="border-top-color:{c};"><div class="sc-val" style="color:{c};">{s}%</div><div class="sc-label">{attr}</div>{sub_div}</div>'
        sc_html += '</div>'


        st.markdown(f'<div class="eyebrow">Relational Quality — % scoring "Quite a bit" or "A great deal"</div>',unsafe_allow_html=True)
        st.markdown(sc_html,unsafe_allow_html=True)

        st.markdown("<br>",unsafe_allow_html=True)
        col_a,col_b=st.columns(2)

        with col_a:
            st.markdown(f'<div class="eyebrow" style="color:{TEAL};">Depth of connection</div>',unsafe_allow_html=True)
            if "Depth of Connection" in edge_df.columns:
                order=["Awareness","Connection","Cooperation","Collaboration"]
                counts=edge_df["Depth of Connection"].value_counts().reindex([o for o in order if o in edge_df["Depth of Connection"].values],fill_value=0)
                fig=go.Figure(go.Bar(x=list(counts.index),y=list(counts.values),
                    marker_color=[SURFACE2,"#c7c5f5",INDIGO_LT,INDIGO],marker_line_width=0,
                    text=list(counts.values),textposition="outside",textfont=dict(family="IBM Plex Mono",size=10)))
                fig.update_layout(paper_bgcolor=SURFACE,plot_bgcolor=SURFACE,showlegend=False,
                    xaxis=dict(tickfont=dict(family="IBM Plex Mono",size=10,color=TEXT2),showgrid=False),
                    yaxis=dict(showgrid=False,showticklabels=False),
                    margin=dict(l=0,r=0,t=16,b=0),height=200)
                st.plotly_chart(fig,use_container_width=True)

            st.markdown(f'<div class="eyebrow" style="color:{TEAL};margin-top:16px;">Reciprocity</div>',unsafe_allow_html=True)
            if "Reciprocity" in edge_df.columns:
                rc=edge_df["Reciprocity"].value_counts()
                fig=go.Figure(go.Bar(x=list(rc.values),y=list(rc.index),orientation="h",
                    marker_color=[GREEN if "balanced" in str(k).lower() else AMBER for k in rc.index],
                    marker_line_width=0,text=list(rc.values),textposition="outside",
                    textfont=dict(family="IBM Plex Mono",size=10)))
                fig.update_layout(paper_bgcolor=SURFACE,plot_bgcolor=SURFACE,showlegend=False,
                    xaxis=dict(showgrid=False,showticklabels=False),
                    yaxis=dict(tickfont=dict(family="IBM Plex Sans",size=11,color=TEXT2),showgrid=False,autorange="reversed"),
                    margin=dict(l=0,r=0,t=8,b=0),height=200)
                st.plotly_chart(fig,use_container_width=True)

        with col_b:
            st.markdown(f'<div class="eyebrow" style="color:{TEAL};">Relationship types</div>',unsafe_allow_html=True)
            if "Relationship Type(s)" in edge_df.columns:
                all_t=[]
                for v in edge_df["Relationship Type(s)"].dropna():
                    all_t.extend([t.strip() for t in str(v).split(";") if t.strip()])
                if all_t:
                    tc=pd.Series(all_t).value_counts()
                    short_labels=[l.split("/")[0].strip() for l in tc.index]
                    fig=go.Figure(go.Bar(x=list(tc.values),y=short_labels,orientation="h",
                        marker_color=TEAL,marker_line_width=0,
                        text=list(tc.values),textposition="outside",textfont=dict(family="IBM Plex Mono",size=10)))
                    fig.update_layout(paper_bgcolor=SURFACE,plot_bgcolor=SURFACE,showlegend=False,
                        xaxis=dict(showgrid=False,showticklabels=False),
                        yaxis=dict(tickfont=dict(family="IBM Plex Sans",size=11,color=TEXT2),showgrid=False,autorange="reversed"),
                        margin=dict(l=0,r=0,t=8,b=0),height=260)
                    st.plotly_chart(fig,use_container_width=True)

            st.markdown(f'<div class="eyebrow" style="color:{TEAL};margin-top:16px;">Frequency</div>',unsafe_allow_html=True)
            if "Frequency of Interaction" in edge_df.columns:
                order=["Never or rarely","Once or twice a month","Weekly","Multiple times a week","Daily or near-daily"]
                fc=edge_df["Frequency of Interaction"].value_counts().reindex([o for o in order if o in edge_df["Frequency of Interaction"].values],fill_value=0)
                fig=go.Figure(go.Bar(x=[o.split(" ")[0] for o in fc.index],y=list(fc.values),
                    marker_color=TEAL,marker_line_width=0,
                    text=list(fc.values),textposition="outside",textfont=dict(family="IBM Plex Mono",size=10)))
                fig.update_layout(paper_bgcolor=SURFACE,plot_bgcolor=SURFACE,showlegend=False,
                    xaxis=dict(tickfont=dict(family="IBM Plex Mono",size=10,color=TEXT2),showgrid=False),
                    yaxis=dict(showgrid=False,showticklabels=False),
                    margin=dict(l=0,r=0,t=8,b=0),height=180)
                st.plotly_chart(fig,use_container_width=True)

# ═══ TAB 3: SKILLS ═══════════════════════════════════════════════════════════
with tab3:
    if pre_df.empty:
        st.info("Upload the pre-survey Excel in the sidebar to enable this tab.")
    else:
        matches=build_skill_match(pre_df)
        st.plotly_chart(skill_gap_chart(matches),use_container_width=True)

        matchable=[m for m in matches if m["offered"]>0 and m["needed"]>0]
        if matchable:
            st.markdown(f'<div class="eyebrow" style="color:{TEAL};margin-top:8px;">Matchmaking opportunities</div>',unsafe_allow_html=True)
            cols=st.columns(2)
            for i,m in enumerate(matchable[:8]):
                with cols[i%2]:
                    o=", ".join(m["offerers"][:3])+("…" if len(m["offerers"])>3 else "")
                    n=", ".join(m["needers"][:3])+("…" if len(m["needers"])>3 else "")
                    st.markdown(f'<div class="skill-match"><div class="sm-skill">{m["skill"]}</div><div class="sm-body"><b style="color:{TEAL};">Offers:</b> {o}<br><b style="color:{AMBER};">Needs:</b> {n}</div></div>',unsafe_allow_html=True)

        if "challenge" in pre_df.columns:
            st.markdown(f'<div class="eyebrow" style="color:{TERRA};margin-top:16px;">Challenges named</div>',unsafe_allow_html=True)
            ch=pre_df[["name","organization","challenge"]].dropna(subset=["challenge"]).rename(columns={"name":"Name","organization":"Organization","challenge":"Challenge"})
            st.dataframe(ch,use_container_width=True,hide_index=True)

# ═══ TAB 4: INSIGHTS ═════════════════════════════════════════════════════════
with tab4:
    if not a:
        st.info("Process data to generate insights.")
    else:
        st.markdown(f'<div class="eyebrow" style="color:{AMBER};">Key findings</div>',unsafe_allow_html=True)
        st.markdown(f'<h2 style="font-family:Barlow Condensed,sans-serif;font-size:32px;text-transform:uppercase;margin-bottom:24px;">What the network shows</h2>',unsafe_allow_html=True)

        # Editorial findings layout
        trust_pct = a.get("pct_trust",0)
        trust_low = a.get("pct_low_trust",0)
        energy_pct = a.get("pct_energy",0)
        energy_low = a.get("pct_low_energy",0)
        trust_body = "Strong foundation." if trust_pct>=60 else f"{trust_low}% of connections score low on trust. These are the relationships most worth attending to."
        energy_body = "Most connections feel generative." if energy_pct>=60 else f"{energy_low}% feel low-energy. Energy tracks closely with trust."
        findings=[
            ("SCALE",f"{a.get('n_nodes',0)} people · {a.get('n_edges',0)} connections",
             f"Network density is {a.get('density',0):.3f}. {'The cohort is well-connected.' if a.get('density',0)>0.15 else 'There is substantial room to build connections across the full group.'}"),
            ("DEPTH",f"{a.get('pct_deep',0)}% at cooperation or collaboration",
             "The remaining connections sit at Awareness or Connection — early but real."),
            ("TRUST",f"{trust_pct}% rate trust highly",
             trust_body),
            ("ENERGY",f"{energy_pct}% feel high-energy",
             energy_body),
            ("RECIPROCITY",f"{a.get('pct_balanced',0)}% described as balanced",
             "Imbalanced flows are normal — mentorship, funding, and support often run one direction. The question is whether people feel it."),
        ]

        if a.get("n_survey_isolated",0)>0:
            findings.append(("ISOLATION",f"{a.get('n_survey_isolated',0)} member(s) not yet named by others",
                "The proposal names leader loneliness as a concern. These are the people most likely to benefit from intentional connection."))

        if a.get("top_bridge"):
            findings.append(("BRIDGE RISK",f"One person carries disproportionate connective load",
                "The person with the highest betweenness centrality sits on more paths between others than anyone else. Their engagement is a structural asset — and a structural vulnerability."))

        if a.get("n_geographies",0)>2:
            findings.append(("GEOGRAPHY",f"Network spans {a.get('n_geographies',0)} areas",
                "Geographic spread is both a strength and a challenge. The clusters that exist likely reflect geography as much as affinity."))

        for lbl_,stat,body in findings:
            color=TERRA if lbl_ in ["ISOLATION","BRIDGE RISK"] else (AMBER if lbl_ in ["TRUST","ENERGY"] and ("low" in body.lower() or "not yet" in body.lower()) else INDIGO)
            st.markdown(f'<div class="finding"><div class="f-label" style="color:{AMBER};">{lbl_}</div><div class="f-stat" style="color:{color};">{stat}</div><div class="f-body">{body}</div></div>',unsafe_allow_html=True)

# ═══ TAB 5: ROOM VIEW ════════════════════════════════════════════════════════
with tab5:
    st.markdown(f'<div class="eyebrow" style="color:{TEAL};">Projectable insights · Activity 5 — Who is in the ecosystem?</div>',unsafe_allow_html=True)
    st.markdown("Preview the anonymized HTML document designed for projection in the room. Download it from the sidebar.")
    st.markdown("<br>",unsafe_allow_html=True)

    if a:
        html_preview=build_insights_html(a,ego_df,edge_df)
        st.components.v1.html(html_preview,height=700,scrolling=True)
    else:
        st.info("Process data to preview the insights document.")
