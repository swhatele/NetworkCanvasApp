"""
NetCanvas Network Analyzer — Connecting for Change
Upload zipped NetCanvas exports, explore the combined network, and download results.
"""

import io, zipfile, xml.etree.ElementTree as ET, math, collections
import pandas as pd
import networkx as nx
import plotly.graph_objects as go
import plotly.express as px
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="TNN Network Analyzer",
    page_icon="🌐",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# STYLE  (C4C palette)
# ─────────────────────────────────────────────
INDIGO   = "#2825BE"
INDIGO_LT= "#4a47d6"
AMBER    = "#EB9001"
TEAL     = "#0C7A7A"
TEAL_LT  = "#0fa8a8"
TERRA    = "#CF4C38"
GREEN    = "#2E9E5B"
INK      = "#080818"
INK2     = "#0f0e2e"
INK3     = "#13123a"
SURFACE2 = "#F5F7F6"
BORDER   = "#e2e5e3"
TEXT     = "#111827"
TEXT2    = "#4b5563"
TEXT3    = "#9ca3af"

NODE_COLORS = [INDIGO, TEAL, AMBER, GREEN, TERRA, INDIGO_LT, TEAL_LT]

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@600;700;800;900&family=IBM+Plex+Sans:ital,wght@0,300;0,400;0,500;0,600;0,700;1,400&family=IBM+Plex+Mono:wght@400;500&display=swap');

html, body, [class*="css"] {{
    font-family: 'IBM Plex Sans', sans-serif;
    color: {TEXT};
}}
h1, h2, h3 {{
    font-family: 'Barlow Condensed', sans-serif;
    letter-spacing: -0.01em;
    color: {INK};
}}
.eyebrow {{
    font-family: 'IBM Plex Sans', sans-serif;
    font-size: 10.5px;
    font-weight: 700;
    letter-spacing: 0.16em;
    text-transform: uppercase;
    color: {AMBER};
    margin-bottom: 4px;
}}
.metric-card {{
    background: {SURFACE2};
    border-radius: 8px;
    padding: 20px 24px;
    border-left: 3px solid {INDIGO};
}}
.metric-card .value {{
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 40px;
    font-weight: 800;
    color: {INDIGO};
    line-height: 1;
}}
.metric-card .label {{
    font-family: 'IBM Plex Mono', monospace;
    font-size: 11px;
    color: {TEXT2};
    margin-top: 4px;
}}
.takeaway-card {{
    background: {INK2};
    border-radius: 8px;
    padding: 20px 24px;
    border-left: 3px solid {AMBER};
    color: white;
    margin-bottom: 12px;
}}
.takeaway-card .tk-label {{
    font-family: 'IBM Plex Mono', monospace;
    font-size: 10px;
    color: {AMBER};
    letter-spacing: 0.12em;
    text-transform: uppercase;
    margin-bottom: 6px;
}}
.takeaway-card .tk-text {{
    font-family: 'IBM Plex Sans', sans-serif;
    font-size: 14px;
    color: rgba(255,255,255,0.85);
    line-height: 1.6;
}}
.section-head {{
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 28px;
    font-weight: 800;
    letter-spacing: -0.01em;
    color: {INK};
    text-transform: uppercase;
    margin: 0 0 4px 0;
}}
.dark-panel {{
    background: {INK};
    border-radius: 10px;
    padding: 32px;
    color: white;
}}
.stTabs [data-baseweb="tab-list"] {{
    gap: 4px;
    border-bottom: 1px solid {BORDER};
}}
.stTabs [data-baseweb="tab"] {{
    font-family: 'IBM Plex Sans', sans-serif;
    font-size: 13px;
    font-weight: 500;
    color: {TEXT2};
    padding: 8px 16px;
    border-radius: 4px 4px 0 0;
}}
.stTabs [aria-selected="true"] {{
    color: {INDIGO} !important;
    border-bottom: 2px solid {INDIGO} !important;
    background: transparent !important;
}}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# LOOKUP TABLES
# ─────────────────────────────────────────────
FREQUENCY_LABELS = {1:"Never or rarely",2:"Once or twice a month",3:"Weekly",4:"Multiple times a week",5:"Daily or near-daily"}
DEPTH_LABELS     = {1:"Awareness",2:"Connection",3:"Cooperation",4:"Collaboration"}
KNOWLEDGE_LABELS = {5:"Not at all",6:"Somewhat",7:"Quite a bit",8:"A great deal"}
SCALE4_LABELS    = {1:"Not at all",2:"Somewhat",3:"Quite a bit",4:"A great deal"}
TYPE_LABELS      = {"Type_1":"Peer Learning / Knowledge Exchange","Type_2":"Joint Programming or Co-delivery","Type_3":"Referrals","Type_4":"Funding","Type_5":"Informal Support / Thought Partnership","Type_6":"Governance or Board-level Connection"}
RECIPROCITY_LABELS={1:"Mostly flows from them to me",2:"Roughly balanced",3:"Mostly flows from me to them",4:"I'm not sure"}
SECTOR_LABELS    = {1:"Education",2:"Health & Human Services",3:"Community Development & Housing",4:"Government & Public Sector",5:"Faith Communities",6:"Business & Commerce",7:"Arts, Culture & Humanities",8:"Civic & Advocacy",9:"Media & Communications",10:"Environment & Conservation",11:"Philanthropy & Funding",12:"Other"}
SHARING_LABELS   = {1:"Place & Built Environment",2:"Recognition",3:"Rhythm & Recurrence",4:"Mutual Obligation",5:"Common Stakes",6:"Institutional Anchors",7:"Story, Memory & Identity"}
CONFIDENTIALITY_LABELS={1:"Yes – share full responses",2:"Yes – aggregate/anonymized only",3:"No – keep private"}
GML_NS = "http://graphml.graphdrawing.org/xmlns"
MAX_BARE = 3

def _gml(t): return f"{{{GML_NS}}}{t}"

def lbl(val, lookup):
    if pd.isna(val): return ""
    try: return lookup.get(int(val), val)
    except: return val

def bool_to_types(row, type_cols):
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
    qs, bare = [], 0
    for _,r in df.iterrows():
        if has_attrs(r): qs.append("Current Connections")
        else:
            bare+=1
            qs.append("Aspired Connections" if bare<=MAX_BARE else ("Informal Influence" if bare<=MAX_BARE*2 else "Unknown"))
    df=df.copy(); df["network_question"]=qs; return df

# ─────────────────────────────────────────────
# PARSERS
# ─────────────────────────────────────────────
def parse_ego(b):
    df=pd.read_csv(io.BytesIO(b))
    if df.empty: return {}
    r=df.iloc[0]
    return {"ego_uuid":r.get("networkCanvasEgoUUID",""),"case_id":r.get("networkCanvasCaseID",""),
            "name":r.get("Name",""),"geography":r.get("Geography",""),
            "sector":lbl(r.get("Sector"),SECTOR_LABELS),"sharing":lbl(r.get("Sharing"),SHARING_LABELS),
            "confidentiality":lbl(r.get("Confidentiality"),CONFIDENTIALITY_LABELS),
            "consent":str(r.get("Consent","")),"expansion":r.get("Expansion",""),
            "open":r.get("Open",""),"tenure":r.get("Tenure",""),
            "session_start":r.get("sessionStart",""),"session_finish":r.get("sessionFinish","")}

def parse_attributes(b, ego_uuid, ego_name):
    df=pd.read_csv(io.BytesIO(b)); df=assign_question(df)
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

def parse_graphml(b, ego_uuid, ego_name, uuid_to_name):
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
        if not zipfile.is_zipfile(io.BytesIO(raw)):
            st.warning(f"⚠️ `{f.name}` is not a valid zip — skipped."); continue
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
    all_egos, all_edges, warnings = [], [], []
    uuid_to_name={}
    for sid, files in buckets.items():
        if "ego" not in files: warnings.append(f"No ego file for `{sid}` — skipped."); continue
        ego=parse_ego(files["ego"])
        ego_uuid=ego.get("ego_uuid",sid)
        ego_name=ego.get("name") or ego.get("case_id") or sid
        uuid_to_name[ego_uuid]=ego_name
        all_egos.append(ego)
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
    return ego_df, edge_df, warnings

# ─────────────────────────────────────────────
# EXCEL BUILDER
# ─────────────────────────────────────────────
def build_excel(ego_df, edge_df):
    buf=io.BytesIO()
    with pd.ExcelWriter(buf,engine="openpyxl") as w:
        ego_df.to_excel(w,sheet_name="Nodes",index=False)
        if not edge_df.empty: edge_df.to_excel(w,sheet_name="Edges",index=False)
    buf.seek(0)
    wb=load_workbook(buf)
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

# ─────────────────────────────────────────────
# NETWORK BUILDER
# ─────────────────────────────────────────────
def build_graph(edge_df):
    G=nx.DiGraph()
    if edge_df.empty: return G
    for _,r in edge_df.iterrows():
        frm=str(r["From"]).strip(); to=str(r["To"]).strip()
        if not frm or not to: continue
        attrs={k:v for k,v in r.items() if k not in ("From","To")}
        if G.has_edge(frm,to):
            G[frm][to]["weight"]=G[frm][to].get("weight",1)+1
        else:
            G.add_edge(frm,to,weight=1,**attrs)
    return G

DEPTH_SCORE={"Awareness":1,"Connection":2,"Cooperation":3,"Collaboration":4}
FREQ_SCORE={"Never or rarely":1,"Once or twice a month":2,"Weekly":3,"Multiple times a week":4,"Daily or near-daily":5}

def edge_strength(row):
    d=DEPTH_SCORE.get(str(row.get("Depth of Connection","")),0)
    f=FREQ_SCORE.get(str(row.get("Frequency of Interaction","")),0)
    return (d+f)/2 if (d or f) else 0

# ─────────────────────────────────────────────
# PLOTLY NETWORK FIGURE
# ─────────────────────────────────────────────
def make_network_fig(G, color_by="degree", highlight_node=None, ego_names=None):
    if len(G.nodes)==0:
        return go.Figure().update_layout(title="No network data")

    UG=G.to_undirected()
    pos=nx.spring_layout(UG, seed=42, k=2.5/math.sqrt(max(len(G.nodes),1)))

    degree=dict(G.degree())
    betweenness=nx.betweenness_centrality(G)
    in_degree=dict(G.in_degree())

    # Node sizes
    max_deg=max(degree.values()) if degree else 1
    node_sizes=[12+28*(degree.get(n,0)/max(max_deg,1)) for n in G.nodes]

    # Node colors
    if color_by=="degree":
        vals=[degree.get(n,0) for n in G.nodes]
        max_v=max(vals) if vals else 1
        node_colors=[f"rgba(40,37,190,{0.4+0.6*(v/max_v)})" for v in vals]
    elif color_by=="betweenness":
        vals=[betweenness.get(n,0) for n in G.nodes]
        max_v=max(vals) if vals else 1
        node_colors=[f"rgba(12,122,122,{0.4+0.6*(v/max_v)})" for v in vals]
    elif color_by=="in_degree":
        vals=[in_degree.get(n,0) for n in G.nodes]
        max_v=max(vals) if vals else 1
        node_colors=[f"rgba(235,144,1,{0.4+0.6*(v/max_v)})" for v in vals]
    elif color_by=="survey_taker":
        ego_set=set(ego_names or [])
        node_colors=[AMBER if n in ego_set else INDIGO for n in G.nodes]
    else:
        node_colors=[INDIGO]*len(G.nodes)

    # Edge traces
    edge_x, edge_y = [], []
    for u,v in G.edges():
        x0,y0=pos[u]; x1,y1=pos[v]
        edge_x+=[x0,x1,None]; edge_y+=[y0,y1,None]

    edge_trace=go.Scatter(x=edge_x,y=edge_y,mode="lines",
        line=dict(width=0.8,color=f"rgba(40,37,190,0.18)"),hoverinfo="none",showlegend=False)

    # Node trace
    node_x=[pos[n][0] for n in G.nodes]
    node_y=[pos[n][1] for n in G.nodes]
    node_text=[f"<b>{n}</b><br>Degree: {degree.get(n,0)}<br>In-degree: {in_degree.get(n,0)}<br>Betweenness: {betweenness.get(n,0):.3f}" for n in G.nodes]

    node_trace=go.Scatter(x=node_x,y=node_y,mode="markers+text",
        text=[n for n in G.nodes],textposition="top center",
        textfont=dict(family="IBM Plex Mono",size=9,color=TEXT2),
        hovertext=node_text,hoverinfo="text",
        marker=dict(size=node_sizes,color=node_colors,
            line=dict(width=1.5,color="white")),
        showlegend=False)

    fig=go.Figure(data=[edge_trace,node_trace])
    fig.update_layout(
        paper_bgcolor=INK2, plot_bgcolor=INK2,
        margin=dict(l=16,r=16,t=16,b=16),
        xaxis=dict(showgrid=False,zeroline=False,showticklabels=False),
        yaxis=dict(showgrid=False,zeroline=False,showticklabels=False),
        height=560,
        font=dict(family="IBM Plex Sans"),
        hoverlabel=dict(bgcolor=INK3,font_color="white",font_family="IBM Plex Sans"),
    )
    return fig

# ─────────────────────────────────────────────
# ANALYTICS HELPERS
# ─────────────────────────────────────────────
def scale4_pct(series, label):
    counts=series.value_counts()
    high=counts.get("A great deal",0)+counts.get("Quite a bit",0)
    total=counts.sum()
    return round(100*high/total) if total else 0

def top_n(series, n=3):
    return series.value_counts().head(n).index.tolist()

def detect_bridge_nodes(G):
    UG=G.to_undirected()
    bw=nx.betweenness_centrality(UG)
    sorted_bw=sorted(bw.items(),key=lambda x:-x[1])
    return sorted_bw[:5]

def generate_takeaways(ego_df, edge_df, G):
    takes=[]
    if edge_df.empty: return takes

    n_nodes=len(G.nodes); n_edges=len(G.edges)
    density=nx.density(G)
    takes.append({"label":"NETWORK SCALE","text":f"{n_nodes} unique people appear across the network with {n_edges} reported connections. Overall network density is {density:.3f}."})

    # Most named people
    if "To" in edge_df.columns:
        top_named=edge_df["To"].value_counts().head(3)
        names=", ".join([f"{n} ({c})" for n,c in top_named.items()])
        takes.append({"label":"MOST NAMED","text":f"The most frequently named connections are: {names}. These people appear across multiple survey takers' networks, suggesting central roles in the broader ecosystem."})

    # Bridge nodes
    bridges=detect_bridge_nodes(G)
    if bridges:
        top_bridge=bridges[0][0]
        takes.append({"label":"BRIDGE NODE","text":f"{top_bridge} has the highest betweenness centrality in the network, meaning they sit on the most paths between other people. Removing them would fragment more connections than any other node."})

    # Depth
    if "Depth of Connection" in edge_df.columns:
        depth_counts=edge_df["Depth of Connection"].value_counts(normalize=True)*100
        collab_pct=depth_counts.get("Collaboration",0)+depth_counts.get("Cooperation",0)
        takes.append({"label":"DEPTH OF CONNECTION","text":f"{collab_pct:.0f}% of reported connections reach Cooperation or Collaboration depth — indicating a network with meaningful working relationships, not just awareness."})

    # Energy / Support / Creativity
    for attr in ["Energy","Support","Creativity","Trust"]:
        if attr in edge_df.columns:
            pct=scale4_pct(edge_df[attr],attr)
            takes.append({"label":attr.upper(),"text":f"{pct}% of connections score 'Quite a bit' or 'A great deal' on {attr}. {'This signals a high-vitality network.' if pct>=60 else 'There may be room to deepen the relational quality of some connections.'}"})

    # Relationship types
    if "Relationship Type(s)" in edge_df.columns:
        all_types=[]
        for val in edge_df["Relationship Type(s)"].dropna():
            all_types.extend([t.strip() for t in str(val).split(";")])
        if all_types:
            top_type=collections.Counter(all_types).most_common(1)[0][0]
            takes.append({"label":"RELATIONSHIP TYPE","text":f"The most common relationship type is '{top_type}'. This characterizes the dominant mode of exchange in the network."})

    # Geography
    if "Geography" in ego_df.columns:
        geos=ego_df["Geography"].dropna()
        if not geos.empty:
            unique_geos=geos.nunique()
            takes.append({"label":"GEOGRAPHY","text":f"Survey takers represent {unique_geos} distinct geographic area(s), suggesting {'a geographically distributed network' if unique_geos>3 else 'geographic concentration — which may reflect either a local focus or a gap in reach'}."})

    # Reciprocity
    if "Reciprocity" in edge_df.columns:
        balanced_pct=round(100*(edge_df["Reciprocity"]=="Roughly balanced").sum()/max(len(edge_df),1))
        takes.append({"label":"RECIPROCITY","text":f"{balanced_pct}% of connections are described as roughly balanced. {'The network has strong mutual exchange.' if balanced_pct>=50 else 'Many connections flow in one direction — worth exploring whether this reflects mentorship structures or unmet reciprocity.'}"})

    return takes

# ─────────────────────────────────────────────
# CHART HELPERS
# ─────────────────────────────────────────────
def bar_chart(labels, values, title, color=INDIGO):
    fig=go.Figure(go.Bar(x=labels,y=values,marker_color=color,marker_line_width=0))
    fig.update_layout(paper_bgcolor="white",plot_bgcolor="white",
        title=dict(text=title,font=dict(family="Barlow Condensed",size=16,color=INK)),
        xaxis=dict(tickfont=dict(family="IBM Plex Mono",size=10,color=TEXT2),showgrid=False,linecolor=BORDER),
        yaxis=dict(tickfont=dict(family="IBM Plex Mono",size=10,color=TEXT2),gridcolor=BORDER,linecolor=BORDER),
        margin=dict(l=16,r=16,t=40,b=16),height=280,font=dict(family="IBM Plex Sans"))
    return fig

def donut_chart(labels, values, title, colors=None):
    if colors is None: colors=[INDIGO,TEAL,AMBER,GREEN,TERRA,INDIGO_LT,TEAL_LT]
    fig=go.Figure(go.Pie(labels=labels,values=values,hole=0.55,
        marker=dict(colors=colors[:len(labels)],line=dict(color="white",width=2)),
        textfont=dict(family="IBM Plex Mono",size=10),showlegend=True))
    fig.update_layout(paper_bgcolor="white",
        title=dict(text=title,font=dict(family="Barlow Condensed",size=16,color=INK)),
        legend=dict(font=dict(family="IBM Plex Sans",size=11,color=TEXT2)),
        margin=dict(l=16,r=16,t=40,b=16),height=280)
    return fig

def attr_bar(edge_df, col, title, order=None, color=TEAL):
    if col not in edge_df.columns: return None
    counts=edge_df[col].value_counts()
    if order: counts=counts.reindex([o for o in order if o in counts.index],fill_value=0)
    return bar_chart(list(counts.index),list(counts.values),title,color)

# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
col_logo, col_title = st.columns([1,6])
with col_title:
    st.markdown('<div class="eyebrow">Connecting for Change</div>', unsafe_allow_html=True)
    st.markdown('<h1 style="font-size:36px;margin:0;">TNN Network Analyzer</h1>', unsafe_allow_html=True)
    st.markdown(f'<p style="color:{TEXT2};font-size:14px;margin-top:4px;">Upload participant exports · Explore the network · Download results</p>', unsafe_allow_html=True)

st.markdown(f'<hr style="border:none;border-top:1px solid {BORDER};margin:16px 0 24px 0">', unsafe_allow_html=True)

# ─────────────────────────────────────────────
# SIDEBAR — UPLOAD
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown('<div class="eyebrow" style="color:#EB9001;">Data Input</div>', unsafe_allow_html=True)
    st.markdown("### Upload exports")
    st.markdown(f'<p style="font-size:13px;color:{TEXT2};">Upload one or more NetCanvas export zip files — one per participant or all together.</p>', unsafe_allow_html=True)

    uploaded = st.file_uploader("Zip files", accept_multiple_files=True, type=["zip"], label_visibility="collapsed")

    if uploaded:
        if st.button("▶  Process", type="primary", use_container_width=True):
            with st.spinner("Processing..."):
                buckets=extract_zips(uploaded)
                ego_df, edge_df, warnings=process_buckets(buckets)
                st.session_state["ego_df"]=ego_df
                st.session_state["edge_df"]=edge_df
                st.session_state["G"]=build_graph(edge_df)
                st.session_state["ego_names"]=list(ego_df["Name"].dropna()) if "Name" in ego_df.columns else []
            for w in warnings: st.warning(w)
            st.success(f"✅ {len(ego_df)} participant(s) · {len(edge_df)} edge(s)")

    if "edge_df" in st.session_state and not st.session_state["edge_df"].empty:
        st.markdown("---")
        st.markdown('<div class="eyebrow" style="color:#EB9001;">Export</div>', unsafe_allow_html=True)
        excel=build_excel(st.session_state["ego_df"],st.session_state["edge_df"])
        st.download_button("⬇  Download Excel",data=excel,file_name="combined_network.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)

# ─────────────────────────────────────────────
# MAIN — requires data
# ─────────────────────────────────────────────
if "ego_df" not in st.session_state:
    st.markdown(f"""
    <div style="background:{SURFACE2};border-radius:10px;padding:48px 32px;text-align:center;border:1px dashed {BORDER};">
        <div style="font-family:'Barlow Condensed',sans-serif;font-size:22px;color:{TEXT2};font-weight:700;text-transform:uppercase;letter-spacing:-0.01em;">Upload zip files in the sidebar to get started</div>
        <div style="font-family:'IBM Plex Sans',sans-serif;font-size:13px;color:{TEXT3};margin-top:8px;">Each participant's export will be combined into a single network</div>
    </div>""", unsafe_allow_html=True)
    st.stop()

ego_df  = st.session_state["ego_df"]
edge_df = st.session_state["edge_df"]
G       = st.session_state["G"]
ego_names = st.session_state.get("ego_names",[])

# ─── Metric strip ───
c1,c2,c3,c4 = st.columns(4)
metrics=[
    (len(ego_df),"SURVEY TAKERS",INDIGO),
    (len(G.nodes),"UNIQUE NODES",TEAL),
    (len(G.edges),"CONNECTIONS",AMBER),
    (f"{nx.density(G):.3f}","NETWORK DENSITY",GREEN),
]
for col,(val,label,color) in zip([c1,c2,c3,c4],metrics):
    with col:
        st.markdown(f"""<div class="metric-card" style="border-left-color:{color};">
            <div class="value" style="color:{color};">{val}</div>
            <div class="label">{label}</div>
        </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ─── Tabs ───
tab1, tab2, tab3, tab4 = st.tabs(["🕸  Network", "📊  Attributes", "📍  Geography", "💡  Takeaways"])

# ═══════════════════════════════════════════
# TAB 1 — NETWORK
# ═══════════════════════════════════════════
with tab1:
    ctrl1, ctrl2 = st.columns([3,1])
    with ctrl2:
        st.markdown("<br>", unsafe_allow_html=True)
        color_by=st.selectbox("Color nodes by",["degree","betweenness","in_degree","survey_taker"],
            format_func=lambda x:{"degree":"Total Degree","betweenness":"Betweenness Centrality","in_degree":"Times Named (In-Degree)","survey_taker":"Survey Taker vs. Named"}.get(x,x))
        show_labels=st.checkbox("Show labels",value=True)

    fig_net=make_network_fig(G, color_by=color_by, ego_names=ego_names)
    if not show_labels:
        fig_net.data[1].mode="markers"
    st.plotly_chart(fig_net,use_container_width=True)

    # Centrality table
    st.markdown(f'<div class="eyebrow" style="color:{TEAL};">Centrality Measures</div>', unsafe_allow_html=True)
    st.markdown("**Top nodes by influence and position**")

    deg=dict(G.degree()); indeg=dict(G.in_degree()); outdeg=dict(G.out_degree())
    bw=nx.betweenness_centrality(G)
    cent_df=pd.DataFrame([{"Name":n,"Degree":deg[n],"Times Named":indeg[n],"Named Others":outdeg[n],"Betweenness":round(bw[n],4)} for n in G.nodes])
    cent_df=cent_df.sort_values("Degree",ascending=False).reset_index(drop=True)
    st.dataframe(cent_df, use_container_width=True, hide_index=True)

    # Community detection
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(f'<div class="eyebrow" style="color:{TEAL};">Community Detection</div>', unsafe_allow_html=True)
    st.markdown("**Clusters detected via Louvain method**")
    UG2=G.to_undirected()
    if len(UG2.nodes)>1:
        try:
            comms=nx.community.louvain_communities(UG2, seed=42)
            comm_rows=[]
            for i,c in enumerate(sorted(comms,key=len,reverse=True)):
                comm_rows.append({"Cluster":f"Cluster {i+1}","Size":len(c),"Members":", ".join(sorted(c))})
            st.dataframe(pd.DataFrame(comm_rows),use_container_width=True,hide_index=True)
        except Exception as e:
            st.info(f"Community detection unavailable: {e}")
    else:
        st.info("Need more nodes for community detection.")

# ═══════════════════════════════════════════
# TAB 2 — ATTRIBUTES
# ═══════════════════════════════════════════
with tab2:
    if edge_df.empty:
        st.info("No edge data available.")
    else:
        col_a, col_b = st.columns(2)

        with col_a:
            f=attr_bar(edge_df,"Depth of Connection","Depth of Connection",
                order=["Awareness","Connection","Cooperation","Collaboration"],color=INDIGO)
            if f: st.plotly_chart(f,use_container_width=True)

            f=attr_bar(edge_df,"Frequency of Interaction","Frequency of Interaction",
                order=["Never or rarely","Once or twice a month","Weekly","Multiple times a week","Daily or near-daily"],color=TEAL)
            if f: st.plotly_chart(f,use_container_width=True)

            f=attr_bar(edge_df,"Reciprocity","Reciprocity",color=AMBER)
            if f: st.plotly_chart(f,use_container_width=True)

        with col_b:
            for attr,color in [("Energy",INDIGO),("Support",TEAL),("Creativity",AMBER),("Trust",GREEN)]:
                f=attr_bar(edge_df,attr,attr,
                    order=["Not at all","Somewhat","Quite a bit","A great deal"],color=color)
                if f: st.plotly_chart(f,use_container_width=True)

        # Relationship types
        st.markdown(f'<div class="eyebrow" style="margin-top:16px;color:{TEAL};">Relationship Types</div>', unsafe_allow_html=True)
        if "Relationship Type(s)" in edge_df.columns:
            all_types=[]
            for val in edge_df["Relationship Type(s)"].dropna():
                all_types.extend([t.strip() for t in str(val).split(";") if t.strip()])
            if all_types:
                type_counts=pd.Series(all_types).value_counts()
                f=bar_chart(list(type_counts.index),list(type_counts.values),"Relationship Types (multi-select)",color=TEAL)
                st.plotly_chart(f,use_container_width=True)

        # Sector (nodes)
        if "Sector" in ego_df.columns:
            st.markdown(f'<div class="eyebrow" style="color:{TEAL};">Sector Distribution (Survey Takers)</div>', unsafe_allow_html=True)
            sc=ego_df["Sector"].value_counts()
            f=donut_chart(list(sc.index),list(sc.values),"Sector")
            st.plotly_chart(f,use_container_width=True)

        # Sharing dimension
        if "Dimension Most Excited to Share" in ego_df.columns:
            st.markdown(f'<div class="eyebrow" style="color:{TEAL};">Dimension Most Excited to Share</div>', unsafe_allow_html=True)
            sh=ego_df["Dimension Most Excited to Share"].value_counts()
            f=bar_chart(list(sh.index),list(sh.values),"Dimension Most Excited to Share",color=INDIGO)
            st.plotly_chart(f,use_container_width=True)

# ═══════════════════════════════════════════
# TAB 3 — GEOGRAPHY
# ═══════════════════════════════════════════
with tab3:
    if "Geography" not in ego_df.columns or ego_df["Geography"].dropna().empty:
        st.info("No geography data available.")
    else:
        geo_counts=ego_df["Geography"].value_counts().reset_index()
        geo_counts.columns=["Geography","Count"]

        fig_geo=go.Figure(go.Bar(
            x=geo_counts["Count"], y=geo_counts["Geography"],
            orientation="h", marker_color=INDIGO, marker_line_width=0))
        fig_geo.update_layout(
            paper_bgcolor="white", plot_bgcolor="white",
            title=dict(text="Survey Takers by Geography",font=dict(family="Barlow Condensed",size=18,color=INK)),
            xaxis=dict(tickfont=dict(family="IBM Plex Mono",size=10,color=TEXT2),gridcolor=BORDER,linecolor=BORDER),
            yaxis=dict(tickfont=dict(family="IBM Plex Sans",size=11,color=TEXT),showgrid=False,linecolor=BORDER,autorange="reversed"),
            margin=dict(l=16,r=16,t=48,b=16), height=max(250,40*len(geo_counts)),
            font=dict(family="IBM Plex Sans"))
        st.plotly_chart(fig_geo,use_container_width=True)

        st.markdown(f'<div class="eyebrow" style="color:{TEAL};">Participant list by location</div>', unsafe_allow_html=True)
        geo_table=ego_df[["Name","Geography","Sector"]].dropna(subset=["Geography"]).sort_values("Geography")
        st.dataframe(geo_table, use_container_width=True, hide_index=True)

# ═══════════════════════════════════════════
# TAB 4 — TAKEAWAYS
# ═══════════════════════════════════════════
with tab4:
    takes=generate_takeaways(ego_df,edge_df,G)
    if not takes:
        st.info("Process data to generate takeaways.")
    else:
        st.markdown(f'<div class="eyebrow" style="color:{AMBER};">Key Findings</div>', unsafe_allow_html=True)
        st.markdown(f'<h2 style="font-family:Barlow Condensed,sans-serif;font-size:28px;text-transform:uppercase;margin-bottom:24px;">What the network shows</h2>', unsafe_allow_html=True)
        for t in takes:
            st.markdown(f"""<div class="takeaway-card">
                <div class="tk-label">{t['label']}</div>
                <div class="tk-text">{t['text']}</div>
            </div>""", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(f'<div class="eyebrow" style="color:{AMBER};">Network Stats</div>', unsafe_allow_html=True)
        st.markdown("**Summary metrics · IBM Plex Mono**")
        if len(G.nodes)>0:
            avg_deg=sum(dict(G.degree()).values())/len(G.nodes)
            try: avg_path=nx.average_shortest_path_length(G.to_undirected()) if nx.is_connected(G.to_undirected()) else "N/A (disconnected)"
            except: avg_path="N/A"
            stats={
                "Nodes":len(G.nodes),"Edges":len(G.edges),
                "Density":f"{nx.density(G):.4f}","Avg Degree":f"{avg_deg:.2f}",
                "Avg Shortest Path":avg_path,
                "Connected Components":nx.number_connected_components(G.to_undirected()),
            }
            stats_df=pd.DataFrame(stats.items(),columns=["Metric","Value"])
            st.dataframe(stats_df, use_container_width=True, hide_index=True)

