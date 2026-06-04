"""
TNN Network Analyzer — Connecting for Change
Upload zipped NetCanvas exports + optional pre-survey Excel to explore the combined network.
"""

import io, zipfile, xml.etree.ElementTree as ET, math, collections
import pandas as pd
import networkx as nx
import plotly.graph_objects as go
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="TNN Network Analyzer", page_icon="🌐",
                   layout="wide", initial_sidebar_state="expanded")

# ── C4C palette ───────────────────────────────────────────────────────────────
INDIGO="#2825BE"; INDIGO_LT="#4a47d6"; AMBER="#EB9001"; TEAL="#0C7A7A"
TEAL_LT="#0fa8a8"; TERRA="#CF4C38"; GREEN="#2E9E5B"; INK="#080818"
INK2="#0f0e2e"; INK3="#13123a"; SURFACE2="#F5F7F6"; BORDER="#e2e5e3"
TEXT="#111827"; TEXT2="#4b5563"; TEXT3="#9ca3af"

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@600;700;800;900&family=IBM+Plex+Sans:ital,wght@0,300;0,400;0,500;0,600;0,700;1,400&family=IBM+Plex+Mono:wght@400;500&display=swap');
html,body,[class*="css"]{{font-family:'IBM Plex Sans',sans-serif;color:{TEXT};}}
h1,h2,h3{{font-family:'Barlow Condensed',sans-serif;letter-spacing:-0.01em;color:{INK};}}
.eyebrow{{font-family:'IBM Plex Sans',sans-serif;font-size:10.5px;font-weight:700;letter-spacing:0.16em;text-transform:uppercase;color:{AMBER};margin-bottom:4px;}}
.metric-card{{background:{SURFACE2};border-radius:8px;padding:20px 24px;border-left:3px solid {INDIGO};}}
.metric-card .value{{font-family:'Barlow Condensed',sans-serif;font-size:40px;font-weight:800;color:{INDIGO};line-height:1;}}
.metric-card .label{{font-family:'IBM Plex Mono',monospace;font-size:11px;color:{TEXT2};margin-top:4px;}}
.takeaway-card{{background:{INK2};border-radius:8px;padding:20px 24px;border-left:3px solid {AMBER};color:white;margin-bottom:12px;}}
.takeaway-card .tk-label{{font-family:'IBM Plex Mono',monospace;font-size:10px;color:{AMBER};letter-spacing:0.12em;text-transform:uppercase;margin-bottom:6px;}}
.takeaway-card .tk-text{{font-family:'IBM Plex Sans',sans-serif;font-size:14px;color:rgba(255,255,255,0.85);line-height:1.6;}}
.skill-match{{background:{INK3};border-radius:6px;padding:12px 16px;margin-bottom:8px;border-left:2px solid {TEAL};}}
.skill-match .sm-skill{{font-family:'IBM Plex Mono',monospace;font-size:11px;color:{TEAL_LT};margin-bottom:4px;}}
.skill-match .sm-body{{font-family:'IBM Plex Sans',sans-serif;font-size:13px;color:rgba(255,255,255,0.8);}}
.stTabs [data-baseweb="tab"]{{font-family:'IBM Plex Sans',sans-serif;font-size:13px;font-weight:500;color:{TEXT2};}}
.stTabs [aria-selected="true"]{{color:{INDIGO}!important;}}
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
DEPTH_SCORE={"Awareness":1,"Connection":2,"Cooperation":3,"Collaboration":4}
FREQ_SCORE={"Never or rarely":1,"Once or twice a month":2,"Weekly":3,"Multiple times a week":4,"Daily or near-daily":5}

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
    df=pd.read_csv(io.BytesIO(b), encoding="utf-8-sig", encoding_errors="replace")
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
    df=pd.read_csv(io.BytesIO(b), encoding="utf-8-sig", encoding_errors="replace"); df=assign_question(df)
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
        edges.append({"From":sn.get("name",""),"To":tn.get("name",""),"ego_uuid":ego_uuid,"ego_name":ego_name})
    return edges

def extract_zips(uploaded_zips):
    buckets={}
    for f in uploaded_zips:
        raw=f.read()
        if not zipfile.is_zipfile(io.BytesIO(raw)): st.warning(f"⚠️ `{f.name}` is not a valid zip — skipped."); continue
        with zipfile.ZipFile(io.BytesIO(raw)) as zf:
            for zname in zf.namelist():
                basename=zname.split("/")[-1]
                if not basename: continue
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

# ── Pre-survey parser ─────────────────────────────────────────────────────────
def parse_presurvey(uploaded_file):
    df=pd.read_excel(uploaded_file,sheet_name="Form Responses 1")
    df=df.rename(columns={
        "Name":"name","Organization Name":"organization",
        "Sector":"sector_raw",
        "Skills you personally can offer (to help others in this network)":"skills_offered",
        "Skills you (or your org) need help with (to accomplish your mission)":"skills_needed",
        "Who is your org's primary target audience? (i.e., who are you serving? The more specific, the better)":"target_audience",
        "How would you describe the most challenging issue you (or your org) face? ":"challenge",
        "How would you describe your context? [Historic / Stagnant]":"ctx_historic",
        "How would you describe your context? [Transitional]":"ctx_transitional",
        "How would you describe your context? [Affluent ]":"ctx_affluent",
        "How would you describe your context? [University / Transient]":"ctx_university",
    })
    # Parse context types into a readable list
    def parse_contexts(row):
        ctx_cols=["ctx_historic","ctx_transitional","ctx_affluent","ctx_university"]
        ctx_labels=["Historic/Stagnant","Transitional","Affluent","University/Transient"]
        found=[]
        for col,lbl_ in zip(ctx_cols,ctx_labels):
            if pd.notna(row.get(col,"")): found.append(lbl_)
        return ", ".join(found) if found else ""
    df["neighborhood_contexts"]=df.apply(parse_contexts,axis=1)
    # Normalize name for fuzzy matching
    df["name_clean"]=df["name"].str.strip().str.lower()
    return df[["name","name_clean","organization","sector_raw","skills_offered","skills_needed","challenge","neighborhood_contexts","target_audience"]]

def merge_presurvey(ego_df, pre_df):
    """Merge pre-survey data onto ego_df by name (case-insensitive)."""
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

# ── Graph builder ─────────────────────────────────────────────────────────────
def build_graph(edge_df):
    G=nx.DiGraph()
    if edge_df.empty: return G
    for _,r in edge_df.iterrows():
        frm=str(r["From"]).strip(); to=str(r["To"]).strip()
        if not frm or not to: continue
        if G.has_edge(frm,to): G[frm][to]["weight"]=G[frm][to].get("weight",1)+1
        else: G.add_edge(frm,to,weight=1)
    return G

# ── Network figure ────────────────────────────────────────────────────────────
def make_network_fig(G, color_by="degree", ego_names=None, node_meta=None, show_labels=True):
    if len(G.nodes)==0: return go.Figure()
    UG=G.to_undirected()
    pos=nx.spring_layout(UG,seed=42,k=2.8/math.sqrt(max(len(G.nodes),1)))
    degree=dict(G.degree()); bw=nx.betweenness_centrality(G); indeg=dict(G.in_degree())
    ego_set=set(ego_names or [])
    max_deg=max(degree.values()) if degree else 1

    node_sizes=[14+32*(degree.get(n,0)/max(max_deg,1)) for n in G.nodes]

    CONTEXT_COLORS={"Historic/Stagnant":TERRA,"Transitional":AMBER,"Affluent":GREEN,"University/Transient":TEAL}

    def get_color(n):
        if color_by=="degree":
            v=degree.get(n,0); return f"rgba(40,37,190,{0.35+0.65*(v/max(max_deg,1))})"
        elif color_by=="betweenness":
            v=bw.get(n,0); mv=max(bw.values()) if bw else 1; return f"rgba(12,122,122,{0.35+0.65*(v/max(mv,0.001))})"
        elif color_by=="in_degree":
            v=indeg.get(n,0); mi=max(indeg.values()) if indeg else 1; return f"rgba(235,144,1,{0.35+0.65*(v/max(mi,1))})"
        elif color_by=="survey_taker":
            return AMBER if n in ego_set else INDIGO
        elif color_by=="isolated":
            return TERRA if indeg.get(n,0)==0 and n not in ego_set else INDIGO
        return INDIGO

    node_colors=[get_color(n) for n in G.nodes]

    # Hover text — enrich with pre-survey if available
    def hover(n):
        lines=[f"<b>{n}</b>","",f"Degree: {degree.get(n,0)}",f"Times named: {indeg.get(n,0)}",f"Betweenness: {bw.get(n,0):.3f}"]
        if node_meta and n in node_meta:
            m=node_meta[n]
            if m.get("challenge"): lines+=["",f"<i>Challenge:</i> {str(m['challenge'])[:120]}"]
            if m.get("skills_offered"): lines+=[f"<i>Offers:</i> {str(m['skills_offered'])[:100]}"]
            if m.get("neighborhood_contexts"): lines+=[f"<i>Context:</i> {m['neighborhood_contexts']}"]
        return "<br>".join(lines)

    node_hover=[hover(n) for n in G.nodes]

    ex,ey=[],[]
    for u,v in G.edges():
        x0,y0=pos[u]; x1,y1=pos[v]
        ex+=[x0,x1,None]; ey+=[y0,y1,None]

    edge_trace=go.Scatter(x=ex,y=ey,mode="lines",line=dict(width=0.8,color=f"rgba(40,37,190,0.18)"),hoverinfo="none",showlegend=False)
    mode="markers+text" if show_labels else "markers"
    node_trace=go.Scatter(
        x=[pos[n][0] for n in G.nodes],y=[pos[n][1] for n in G.nodes],
        mode=mode,text=[n for n in G.nodes],textposition="top center",
        textfont=dict(family="IBM Plex Mono",size=9,color=TEXT2),
        hovertext=node_hover,hoverinfo="text",
        marker=dict(size=node_sizes,color=node_colors,line=dict(width=1.5,color="white")),
        showlegend=False)

    fig=go.Figure(data=[edge_trace,node_trace])
    fig.update_layout(paper_bgcolor=INK2,plot_bgcolor=INK2,margin=dict(l=16,r=16,t=16,b=16),
        xaxis=dict(showgrid=False,zeroline=False,showticklabels=False),
        yaxis=dict(showgrid=False,zeroline=False,showticklabels=False),
        height=580,font=dict(family="IBM Plex Sans"),
        hoverlabel=dict(bgcolor=INK3,font_color="white",font_family="IBM Plex Sans",align="left"))
    return fig

# ── Chart helpers ─────────────────────────────────────────────────────────────
def bar_chart(labels,values,title,color=INDIGO,height=280):
    fig=go.Figure(go.Bar(x=labels,y=values,marker_color=color,marker_line_width=0))
    fig.update_layout(paper_bgcolor="white",plot_bgcolor="white",
        title=dict(text=title,font=dict(family="Barlow Condensed",size=16,color=INK)),
        xaxis=dict(tickfont=dict(family="IBM Plex Mono",size=10,color=TEXT2),showgrid=False,linecolor=BORDER),
        yaxis=dict(tickfont=dict(family="IBM Plex Mono",size=10,color=TEXT2),gridcolor=BORDER,linecolor=BORDER),
        margin=dict(l=16,r=16,t=40,b=80),height=height,font=dict(family="IBM Plex Sans"))
    return fig

def hbar_chart(labels,values,title,color=INDIGO,height=None):
    h=height or max(240,30*len(labels)+60)
    fig=go.Figure(go.Bar(x=values,y=labels,orientation="h",marker_color=color,marker_line_width=0))
    fig.update_layout(paper_bgcolor="white",plot_bgcolor="white",
        title=dict(text=title,font=dict(family="Barlow Condensed",size=16,color=INK)),
        xaxis=dict(tickfont=dict(family="IBM Plex Mono",size=10,color=TEXT2),gridcolor=BORDER,linecolor=BORDER),
        yaxis=dict(tickfont=dict(family="IBM Plex Sans",size=11,color=TEXT),showgrid=False,linecolor=BORDER,autorange="reversed"),
        margin=dict(l=16,r=16,t=40,b=16),height=h,font=dict(family="IBM Plex Sans"))
    return fig

def attr_bar(edge_df,col,title,order=None,color=TEAL):
    if col not in edge_df.columns: return None
    counts=edge_df[col].value_counts()
    if order: counts=counts.reindex([o for o in order if o in counts.index],fill_value=0)
    return bar_chart(list(counts.index),list(counts.values),title,color)

# ── Skill analysis ────────────────────────────────────────────────────────────
def parse_skills(skill_str):
    if pd.isna(skill_str) or not str(skill_str).strip(): return []
    return [s.strip() for s in str(skill_str).split(",") if s.strip()]

def build_skill_match(pre_df, ego_names_set):
    """Return list of {skill, needers, offerers} where there's a match."""
    offerers=collections.defaultdict(list)
    needers=collections.defaultdict(list)
    for _,r in pre_df.iterrows():
        name=r.get("name","")
        for sk in parse_skills(r.get("skills_offered","")): offerers[sk].append(name)
        for sk in parse_skills(r.get("skills_needed","")): needers[sk].append(name)
    matches=[]
    all_skills=set(offerers.keys())|set(needers.keys())
    for sk in sorted(all_skills):
        o=offerers.get(sk,[]); n=needers.get(sk,[])
        gap=len(n)-len(o)
        matches.append({"skill":sk,"offered":len(o),"needed":len(n),"gap":gap,
                        "offerers":o,"needers":n})
    return sorted(matches,key=lambda x:-x["needed"])

def skill_gap_chart(matches):
    skills=[m["skill"] for m in matches[:12]]
    offered=[m["offered"] for m in matches[:12]]
    needed=[m["needed"] for m in matches[:12]]
    fig=go.Figure()
    fig.add_trace(go.Bar(name="Offered",x=skills,y=offered,marker_color=TEAL,marker_line_width=0))
    fig.add_trace(go.Bar(name="Needed",x=skills,y=needed,marker_color=AMBER,marker_line_width=0))
    fig.update_layout(barmode="group",paper_bgcolor="white",plot_bgcolor="white",
        title=dict(text="Skills Offered vs. Needed",font=dict(family="Barlow Condensed",size=16,color=INK)),
        xaxis=dict(tickfont=dict(family="IBM Plex Mono",size=9,color=TEXT2),showgrid=False,linecolor=BORDER,tickangle=-30),
        yaxis=dict(tickfont=dict(family="IBM Plex Mono",size=10,color=TEXT2),gridcolor=BORDER,linecolor=BORDER),
        legend=dict(font=dict(family="IBM Plex Sans",size=11)),
        margin=dict(l=16,r=16,t=40,b=100),height=320,font=dict(family="IBM Plex Sans"))
    return fig

# ── Takeaways ─────────────────────────────────────────────────────────────────
def generate_takeaways(ego_df,edge_df,G,pre_df=None):
    takes=[]
    if edge_df.empty: return takes
    n_nodes=len(G.nodes); n_edges=len(G.edges); density=nx.density(G)
    takes.append({"label":"NETWORK SCALE","text":f"{n_nodes} unique people appear across the network with {n_edges} reported connections. Overall network density is {density:.3f} — {'relatively dense, suggesting a well-connected cohort.' if density>0.15 else 'relatively sparse, suggesting room to deepen connections.'}"})

    # Most named
    if "To" in edge_df.columns:
        top=edge_df["To"].value_counts().head(3)
        names=", ".join([f"{n} ({c})" for n,c in top.items()])
        takes.append({"label":"MOST NAMED","text":f"The most frequently named connections are: {names}. These people are central to the network's informal infrastructure."})

    # Isolated nodes (in-degree 0, not a survey taker)
    ego_set=set(ego_df["Name"].dropna().str.strip()) if "Name" in ego_df.columns else set()
    indeg=dict(G.in_degree())
    isolated=[n for n in G.nodes if indeg.get(n,0)==0 and n not in ego_set]
    survey_lonely=[n for n in ego_set if indeg.get(n,0)==0]
    if survey_lonely:
        takes.append({"label":"POTENTIAL ISOLATION","text":f"The following survey taker(s) were not named by anyone else in the network: {', '.join(sorted(survey_lonely))}. The proposal explicitly names leader loneliness as a concern — these individuals may benefit from intentional bridge-building."})
    elif isolated:
        takes.append({"label":"PERIPHERY NODES","text":f"{len(isolated)} named node(s) appear in only one person's network and were not named by anyone else. These may be important connections not yet shared across the cohort."})

    # Bridge nodes
    bw=nx.betweenness_centrality(G)
    top_bridge=sorted(bw.items(),key=lambda x:-x[1])
    if top_bridge:
        name,score=top_bridge[0]
        takes.append({"label":"BRIDGE NODE","text":f"{name} has the highest betweenness centrality ({score:.3f}), sitting on more paths between other people than anyone else. They are a key connector — and a structural vulnerability if they were to disengage."})

    # Depth
    if "Depth of Connection" in edge_df.columns:
        dc=edge_df["Depth of Connection"].value_counts(normalize=True)*100
        deep=dc.get("Collaboration",0)+dc.get("Cooperation",0)
        takes.append({"label":"DEPTH OF CONNECTION","text":f"{deep:.0f}% of connections reach Cooperation or Collaboration depth. {'The cohort has substantive working relationships at its core.' if deep>=50 else 'Most connections are still at the Awareness or Connection stage — depth-building is an opportunity.'}"})

    # Relational quality
    for attr in ["Energy","Support","Creativity","Trust"]:
        if attr in edge_df.columns:
            counts=edge_df[attr].value_counts()
            high=counts.get("A great deal",0)+counts.get("Quite a bit",0)
            pct=round(100*high/max(counts.sum(),1))
            takes.append({"label":attr.upper(),"text":f"{pct}% of connections score 'Quite a bit' or 'A great deal' on {attr}. {'A healthy signal.' if pct>=60 else 'Worth exploring which relationships could be better resourced.'}"})

    # Reciprocity
    if "Reciprocity" in edge_df.columns:
        bal=round(100*(edge_df["Reciprocity"]=="Roughly balanced").sum()/max(len(edge_df),1))
    takes.append({"label":"RECIPROCITY","text":f"{bal}% of connections are described as roughly balanced. " + ("Strong mutual exchange across the network." if bal>=50 else "Many connections flow primarily in one direction — the proposal goal of mutual care and reciprocity has room to develop.")})

    # Relationship types
    if "Relationship Type(s)" in edge_df.columns:
        all_t=[]
        for v in edge_df["Relationship Type(s)"].dropna():
            all_t.extend([t.strip() for t in str(v).split(";") if t.strip()])
        if all_t:
            top_t=collections.Counter(all_t).most_common(1)[0][0]
            funding_ct=sum(1 for t in all_t if "Funding" in t)
            takes.append({"label":"RELATIONSHIP TYPE","text":f"The dominant relationship type is '{top_t}'. {f'Funding relationships account for {funding_ct} connections — worth noting alongside reciprocity patterns.' if funding_ct>0 else ''}"})

    # Skill gap (from pre-survey)
    if pre_df is not None and not pre_df.empty:
        matches=build_skill_match(pre_df,ego_set)
        top_gap=[m for m in matches if m["gap"]>0][:2]
        if top_gap:
            gap_strs=[f"{m['skill']} (needed by {m['needed']}, offered by {m['offered']})" for m in top_gap]
            takes.append({"label":"SKILL GAP","text":f"The largest unmet skill needs in the network are: {'; '.join(gap_strs)}. The proposal calls for smarter resource-sharing — these gaps are the most concrete place to start."})

        # Skills where supply exceeds demand
        surplus=[m for m in matches if m["gap"]<0][:2]
        if surplus:
            sur_strs=[f"{m['skill']} ({m['offered']} offering, {m['needed']} needing)" for m in surplus]
            takes.append({"label":"SKILL SURPLUS","text":f"Skills with more supply than demand: {'; '.join(sur_strs)}. These represent underutilized capacity that could be shared more intentionally."})

    # Context diversity
    if pre_df is not None and "neighborhood_contexts" in pre_df.columns:
        all_ctx=[]
        for v in pre_df["neighborhood_contexts"].dropna():
            all_ctx.extend([c.strip() for c in str(v).split(",") if c.strip()])
        if all_ctx:
            ctx_counts=collections.Counter(all_ctx)
            top_ctx=ctx_counts.most_common(1)[0]
            takes.append({"label":"NEIGHBORHOOD CONTEXT","text":f"The most common neighborhood context is '{top_ctx[0]}' ({top_ctx[1]} organizations). The cohort spans {len(ctx_counts)} context types — a sign of diverse geographic and social experience that enriches peer learning."})

    return takes

# ─────────────────────────────────────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(f'<div class="eyebrow">Connecting for Change</div>', unsafe_allow_html=True)
st.markdown('<h1 style="font-size:38px;margin:0 0 4px 0;">TNN Network Analyzer</h1>', unsafe_allow_html=True)
st.markdown(f'<p style="color:{TEXT2};font-size:14px;margin:0 0 16px 0;">Upload participant exports · Explore the network · Download results</p>', unsafe_allow_html=True)
st.markdown(f'<hr style="border:none;border-top:1px solid {BORDER};margin:0 0 24px 0">', unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('<div class="eyebrow">Data Input</div>', unsafe_allow_html=True)
    st.markdown("### NetCanvas exports")
    uploaded=st.file_uploader("Zip files",accept_multiple_files=True,type=["zip"],label_visibility="collapsed")

    st.markdown("---")
    st.markdown("### Pre-survey responses")
    st.markdown(f'<p style="font-size:12px;color:{TEXT2};">Optional — upload the Network Activation Survey Excel to enrich the analysis with skills, challenges, and neighborhood contexts.</p>', unsafe_allow_html=True)
    presurvey_file=st.file_uploader("Excel file",type=["xlsx"],label_visibility="collapsed",key="presurvey")

    if uploaded:
        if st.button("▶  Process", type="primary", use_container_width=True):
            with st.spinner("Processing..."):
                buckets=extract_zips(uploaded)
                ego_df,edge_df,warnings=process_buckets(buckets)
                pre_df=parse_presurvey(presurvey_file) if presurvey_file else pd.DataFrame()
                if not pre_df.empty: ego_df=merge_presurvey(ego_df,pre_df)
                G=build_graph(edge_df)
                ego_names=list(ego_df["Name"].dropna()) if "Name" in ego_df.columns else []
                # Build node_meta lookup for hover enrichment
                node_meta={}
                if not pre_df.empty:
                    for _,r in pre_df.iterrows():
                        node_meta[r["name"]]={"challenge":r.get("challenge",""),"skills_offered":r.get("skills_offered",""),"neighborhood_contexts":r.get("neighborhood_contexts","")}
                st.session_state.update({"ego_df":ego_df,"edge_df":edge_df,"G":G,"ego_names":ego_names,"pre_df":pre_df,"node_meta":node_meta})
            for w in warnings: st.warning(w)
            st.success(f"✅ {len(ego_df)} participant(s) · {len(edge_df)} edge(s)")

    if "edge_df" in st.session_state and not st.session_state["edge_df"].empty:
        st.markdown("---")
        st.markdown('<div class="eyebrow">Export</div>', unsafe_allow_html=True)
        excel=build_excel(st.session_state["ego_df"],st.session_state["edge_df"])
        st.download_button("⬇  Download Excel",data=excel,file_name="combined_network.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)

# ── Gate ──────────────────────────────────────────────────────────────────────
if "ego_df" not in st.session_state:
    st.markdown(f"""<div style="background:{SURFACE2};border-radius:10px;padding:48px 32px;text-align:center;border:1px dashed {BORDER};">
        <div style="font-family:'Barlow Condensed',sans-serif;font-size:22px;color:{TEXT2};font-weight:700;text-transform:uppercase;">Upload zip files in the sidebar to get started</div>
        <div style="font-family:'IBM Plex Sans',sans-serif;font-size:13px;color:{TEXT3};margin-top:8px;">The pre-survey Excel is optional but enriches every panel with skills, challenges, and neighborhood context</div>
    </div>""", unsafe_allow_html=True)
    st.stop()

ego_df=st.session_state["ego_df"]; edge_df=st.session_state["edge_df"]
G=st.session_state["G"]; ego_names=st.session_state.get("ego_names",[])
pre_df=st.session_state.get("pre_df",pd.DataFrame()); node_meta=st.session_state.get("node_meta",{})

# ── Metric strip ──────────────────────────────────────────────────────────────
c1,c2,c3,c4=st.columns(4)
for col,(val,lbl_,color) in zip([c1,c2,c3,c4],[
    (len(ego_df),"SURVEY TAKERS",INDIGO),(len(G.nodes),"UNIQUE NODES",TEAL),
    (len(G.edges),"CONNECTIONS",AMBER),(f"{nx.density(G):.3f}","NETWORK DENSITY",GREEN)]):
    with col:
        st.markdown(f'<div class="metric-card" style="border-left-color:{color};"><div class="value" style="color:{color};">{val}</div><div class="label">{lbl_}</div></div>',unsafe_allow_html=True)

st.markdown("<br>",unsafe_allow_html=True)

# ── Tabs ──────────────────────────────────────────────────────────────────────
tabs=["🕸  Network","📊  Attributes","🛠  Skills & Challenges","📍  Geography","💡  Takeaways"]
tab1,tab2,tab3,tab4,tab5=st.tabs(tabs)

# ═══ TAB 1: NETWORK ══════════════════════════════════════════════════════════
with tab1:
    c_graph,c_ctrl=st.columns([4,1])
    with c_ctrl:
        st.markdown("<br>",unsafe_allow_html=True)
        color_by=st.selectbox("Color by",["degree","betweenness","in_degree","survey_taker","isolated"],
            format_func=lambda x:{"degree":"Total Degree","betweenness":"Betweenness","in_degree":"Times Named","survey_taker":"Survey Taker vs. Named","isolated":"Potential Isolation"}.get(x,x))
        show_labels=st.checkbox("Labels",value=True)
        st.markdown(f'<p style="font-size:11px;color:{TEXT3};font-family:IBM Plex Mono,monospace;">Node size = degree<br>Hover for details</p>',unsafe_allow_html=True)

    with c_graph:
        fig=make_network_fig(G,color_by=color_by,ego_names=ego_names,node_meta=node_meta,show_labels=show_labels)
        st.plotly_chart(fig,use_container_width=True)

    st.markdown(f'<div class="eyebrow" style="color:{TEAL};">Centrality Measures</div>',unsafe_allow_html=True)
    deg=dict(G.degree()); indeg=dict(G.in_degree()); outdeg=dict(G.out_degree()); bw=nx.betweenness_centrality(G)
    cent_df=pd.DataFrame([{"Name":n,"Degree":deg[n],"Times Named (In)":indeg[n],"Named Others (Out)":outdeg[n],"Betweenness":round(bw[n],4)} for n in G.nodes])
    cent_df=cent_df.sort_values("Degree",ascending=False).reset_index(drop=True)
    st.dataframe(cent_df,use_container_width=True,hide_index=True)

    st.markdown("<br>",unsafe_allow_html=True)
    st.markdown(f'<div class="eyebrow" style="color:{TEAL};">Community Detection · Louvain</div>',unsafe_allow_html=True)
    UG2=G.to_undirected()
    if len(UG2.nodes)>1:
        try:
            comms=nx.community.louvain_communities(UG2,seed=42)
            rows=[{"Cluster":f"Cluster {i+1}","Size":len(c),"Members":", ".join(sorted(c))} for i,c in enumerate(sorted(comms,key=len,reverse=True))]
            st.dataframe(pd.DataFrame(rows),use_container_width=True,hide_index=True)
        except Exception as e: st.info(f"Community detection unavailable with current data size.")
    else: st.info("Need more nodes for community detection.")

# ═══ TAB 2: ATTRIBUTES ═══════════════════════════════════════════════════════
with tab2:
    if edge_df.empty: st.info("No edge data available.")
    else:
        ca1,ca2=st.columns(2)
        with ca1:
            for col,order,color in [
                ("Depth of Connection",["Awareness","Connection","Cooperation","Collaboration"],INDIGO),
                ("Frequency of Interaction",["Never or rarely","Once or twice a month","Weekly","Multiple times a week","Daily or near-daily"],TEAL),
                ("Reciprocity",None,AMBER)]:
                f=attr_bar(edge_df,col,col,order,color)
                if f: st.plotly_chart(f,use_container_width=True)
        with ca2:
            for attr,color in [("Energy",INDIGO),("Support",TEAL),("Creativity",AMBER),("Trust",GREEN)]:
                f=attr_bar(edge_df,attr,attr,["Not at all","Somewhat","Quite a bit","A great deal"],color)
                if f: st.plotly_chart(f,use_container_width=True)

        st.markdown(f'<div class="eyebrow" style="color:{TEAL};">Relationship Types</div>',unsafe_allow_html=True)
        if "Relationship Type(s)" in edge_df.columns:
            all_t=[]
            for v in edge_df["Relationship Type(s)"].dropna():
                all_t.extend([t.strip() for t in str(v).split(";") if t.strip()])
            if all_t:
                tc=pd.Series(all_t).value_counts()
                st.plotly_chart(hbar_chart(list(tc.index),list(tc.values),"Relationship Types (multi-select)",TEAL),use_container_width=True)

        if "Sector" in ego_df.columns:
            sc=ego_df["Sector"].value_counts()
            st.plotly_chart(hbar_chart(list(sc.index),list(sc.values),"Sector Distribution (Survey Takers)",INDIGO),use_container_width=True)

        if "Dimension Most Excited to Share" in ego_df.columns:
            sh=ego_df["Dimension Most Excited to Share"].value_counts()
            st.plotly_chart(bar_chart(list(sh.index),list(sh.values),"Dimension Most Excited to Share",INDIGO),use_container_width=True)

# ═══ TAB 3: SKILLS & CHALLENGES ══════════════════════════════════════════════
with tab3:
    if pre_df.empty:
        st.info("Upload the pre-survey Excel in the sidebar to enable this tab.")
    else:
        matches=build_skill_match(pre_df,set(ego_names))
        st.markdown(f'<div class="eyebrow" style="color:{TEAL};">Skills Offered vs. Needed</div>',unsafe_allow_html=True)
        st.plotly_chart(skill_gap_chart(matches),use_container_width=True)

        st.markdown("<br>",unsafe_allow_html=True)
        st.markdown(f'<div class="eyebrow" style="color:{AMBER};">Skill Matchmaking</div>',unsafe_allow_html=True)
        st.markdown("People who can help others with what they need most.")

        # Only show skills where there's both supply and demand
        matchable=[m for m in matches if m["offered"]>0 and m["needed"]>0]
        cols_sk=st.columns(2)
        for i,m in enumerate(matchable[:10]):
            with cols_sk[i%2]:
                offerer_str=", ".join(m["offerers"][:4])+("…" if len(m["offerers"])>4 else "")
                needer_str=", ".join(m["needers"][:4])+("…" if len(m["needers"])>4 else "")
                st.markdown(f"""<div class="skill-match">
                    <div class="sm-skill">{m['skill']}</div>
                    <div class="sm-body">🟢 Offers: {offerer_str}<br>🟡 Needs: {needer_str}</div>
                </div>""",unsafe_allow_html=True)

        st.markdown("<br>",unsafe_allow_html=True)
        st.markdown(f'<div class="eyebrow" style="color:{TERRA};">Challenges Named</div>',unsafe_allow_html=True)
        st.markdown("What each person or organization identified as their hardest problem.")
        if "challenge" in pre_df.columns:
            ch_df=pre_df[["name","organization","challenge"]].dropna(subset=["challenge"]).rename(columns={"name":"Name","organization":"Organization","challenge":"Challenge"})
            st.dataframe(ch_df,use_container_width=True,hide_index=True)

        st.markdown("<br>",unsafe_allow_html=True)
        st.markdown(f'<div class="eyebrow" style="color:{TEAL};">Neighborhood Contexts</div>',unsafe_allow_html=True)
        all_ctx=[]
        for v in pre_df["neighborhood_contexts"].dropna():
            all_ctx.extend([c.strip() for c in str(v).split(",") if c.strip()])
        if all_ctx:
            ctx_counts=pd.Series(all_ctx).value_counts()
            st.plotly_chart(hbar_chart(list(ctx_counts.index),list(ctx_counts.values),"Neighborhood Contexts (multi-select)",TEAL),use_container_width=True)

# ═══ TAB 4: GEOGRAPHY ════════════════════════════════════════════════════════
with tab4:
    if "Geography" not in ego_df.columns or ego_df["Geography"].dropna().empty:
        st.info("No geography data available from the network survey.")
    else:
        geo_counts=ego_df["Geography"].value_counts()
        st.plotly_chart(hbar_chart(list(geo_counts.index),list(geo_counts.values),"Survey Takers by Geography",INDIGO),use_container_width=True)
        cols_geo=["Name","Geography","Sector"]
        if "skills_offered" in ego_df.columns: cols_geo.append("skills_offered")
        geo_table=ego_df[[c for c in cols_geo if c in ego_df.columns]].dropna(subset=["Geography"]).sort_values("Geography")
        st.dataframe(geo_table,use_container_width=True,hide_index=True)

# ═══ TAB 5: TAKEAWAYS ════════════════════════════════════════════════════════
with tab5:
    takes=generate_takeaways(ego_df,edge_df,G,pre_df if not pre_df.empty else None)
    if not takes: st.info("Process data to generate takeaways.")
    else:
        st.markdown(f'<div class="eyebrow" style="color:{AMBER};">Key Findings</div>',unsafe_allow_html=True)
        st.markdown(f'<h2 style="font-family:Barlow Condensed,sans-serif;font-size:28px;text-transform:uppercase;margin-bottom:24px;">What the network shows</h2>',unsafe_allow_html=True)
        for t in takes:
            st.markdown(f'<div class="takeaway-card"><div class="tk-label">{t["label"]}</div><div class="tk-text">{t["text"]}</div></div>',unsafe_allow_html=True)

        st.markdown("<br>",unsafe_allow_html=True)
        st.markdown(f'<div class="eyebrow" style="color:{AMBER};">Network Stats</div>',unsafe_allow_html=True)
        if len(G.nodes)>0:
            avg_deg=sum(dict(G.degree()).values())/len(G.nodes)
            try: avg_path=nx.average_shortest_path_length(G.to_undirected()) if nx.is_connected(G.to_undirected()) else "N/A (disconnected)"
            except: avg_path="N/A"
            stats_df=pd.DataFrame({"Metric":["Nodes","Edges","Density","Avg Degree","Avg Shortest Path","Connected Components"],
                "Value":[len(G.nodes),len(G.edges),f"{nx.density(G):.4f}",f"{avg_deg:.2f}",avg_path,nx.number_connected_components(G.to_undirected())]})
            st.dataframe(stats_df,use_container_width=True,hide_index=True)
