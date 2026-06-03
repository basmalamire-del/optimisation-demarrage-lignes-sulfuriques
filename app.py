# -*- coding: utf-8 -*-
from __future__ import annotations
import streamlit.components.v1 as components
import json
from pathlib import Path
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import streamlit as st

BASE = Path(__file__).resolve().parent
DATA = BASE / 'data'
ANCIEN = DATA / 'ancien.xlsx'
NOUVEAU = DATA / 'nouveau.xlsx'
RESULTATS = DATA / 'resultats_demarrage.json'

st.set_page_config(page_title='OCP Safi | Préchauffage Convertisseur', page_icon='🟢', layout='wide')

OCP = '#0AA35C'
NAVY = '#052B4F'
NAVY2 = '#073A6B'
LIGHT = '#F5F8FB'
CARD = '#FFFFFF'
TEXT = '#17212B'
MUTED = '#667085'
RED = '#EF4444'
BLUE = '#2563EB'
ORANGE = '#F97316'
YELLOW = '#EAB308'
GREEN = '#22C55E'
GRID = 'rgba(2, 44, 77, .12)'

st.markdown(f'''
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap');
html, body, [class*="css"] {{font-family: Inter, Arial, sans-serif;}}
.stApp {{background:{LIGHT}; color:{TEXT};}}
.block-container {{padding-top: 2.8rem !important; margin-top: 0px !important; padding-bottom:1rem; max-width:1500px;}}
[data-testid="stSidebar"] {{background:linear-gradient(180deg,{NAVY} 0%, #031A31 100%);}}
[data-testid="stSidebar"] * {{color:white !important;}}
[data-testid="stSidebar"] .stRadio label {{background:rgba(255,255,255,.07); border-radius:10px; padding:6px 10px; margin:2px 0;}}
h1,h2,h3 {{color:{TEXT};}}
.hero {{background:{CARD}; border:1px solid #E7ECF2; border-radius:18px; padding:20px 24px; box-shadow:0 8px 24px rgba(16,24,40,.06);}}
.hero h1 {{font-size:26px; margin:0; font-weight:800; color:{NAVY};}}
.hero p {{margin:6px 0 0 0; color:{MUTED};}}
/* Élimination des espaces vides inutiles à gauche et à droite de l'écran */
.block-container {{
    padding-top: 2.8rem !important;
    margin-top: 0px !important;
    padding-left: 1rem !important;
    padding-right: 1rem !important;
    max-width: 98% !important;
}}

/* Configuration stricte des KPI sur une seule ligne */
.kpi {{
    background: #FFFFFF; 
    border: 1px solid #E7ECF2; 
    border-radius: 12px; 
    padding: 10px 12px; 
    height: 90px; 
    box-shadow: 0 4px 12px rgba(16,24,40,.04); 
    display: flex; 
    align-items: center; 
    gap: 10px;
    position: relative;
    overflow: hidden;
}}
.kpi-icon {{
    font-size: 26px; 
    display: flex;
    align-items: center;
    justify-content: center;
    min-width: 32px;
}}
.kpi-content {{
    display: flex;
    flex-direction: column;
    justify-content: center;
    flex-grow: 1;
    overflow: hidden;
}}
.kpi .label {{
    font-size: 10px; 
    text-transform: uppercase; 
    letter-spacing: .03em; 
    color: #667085; 
    font-weight: 700;
    margin: 0;
    line-height: 1.2;
    white-space: nowrap;
}}
.kpi .value {{
    font-size: 18px; 
    color: #052B4F; 
    font-weight: 800; 
    margin: 2px 0;
    line-height: 1.2;
    white-space: nowrap; /* Interdit strictement le retour à la ligne */
}}
.kpi .unit {{
    font-size: 10px; 
    color: #667085; 
    font-weight: 500;
    margin: 0;
    line-height: 1.2;
    white-space: nowrap;
}}
.card {{background:{CARD}; border:1px solid #E7ECF2; border-radius:16px; padding:18px; box-shadow:0 8px 20px rgba(16,24,40,.06);}}
.small-title {{font-size:14px; font-weight:800; color:{NAVY}; text-transform:uppercase; letter-spacing:.04em; margin-bottom:8px;}}
.badge-on {{display:inline-block; padding:7px 11px; border-radius:999px; background:#E9F9F1; color:#067647; font-weight:800; font-size:12px; border:1px solid #A6E8C2;}}
.badge-off {{display:inline-block; padding:7px 11px; border-radius:999px; background:#FEF3F2; color:#B42318; font-weight:800; font-size:12px; border:1px solid #FECDCA;}}
.footer {{color:#98A2B3; font-size:12px; text-align:right; margin-top:8px;}}
hr {{border:0; border-top:1px solid #E7ECF2;}}
</style>
''', unsafe_allow_html=True)

@st.cache_data(show_spinner=False)
def load_all():
    ancien = pd.read_excel(ANCIEN, sheet_name=None, engine="openpyxl")
    nouveau = pd.read_excel(NOUVEAU, sheet_name=None, engine="openpyxl")

    with open(RESULTATS, "r", encoding="utf-8") as f:
        opt = json.load(f)

    for pack in (ancien, nouveau):
        for k in pack:
            pack[k].columns = [str(c).strip() for c in pack[k].columns]

    return ancien, nouveau, opt


try:
    ancien, nouveau, opt = load_all()

    anc_m = ancien["02_Masses"]
    anc_four = ancien["03_Four"]
    anc_gas = ancien["04_Gasoil"]
    anc_kpi = ancien["05_KPI"]
    anc_eco = ancien.get("06_Economie", pd.DataFrame())

    new_m = nouveau["02_Masses"]
    new_four = nouveau["03_Four"]
    new_va = nouveau["04_Vapeur_Air"]
    new_gas = nouveau["05_Gasoil"]
    new_kpi = nouveau["06_KPI"]
    new_eco = nouveau.get("07_Economie", pd.DataFrame())

except Exception as e:
    st.error("Erreur de chargement des fichiers Excel ou JSON.")
    st.exception(e)
    st.stop()

T_ARRET = 11 + 28/60
T_TARGET = 120

# Variables d'icônes pour correspondre à la maquette
# Dictionnaire des icônes émojis épurées
ICONS = {
    'time': '⏱️',
    'gasoil': '💧',
    'h2so4': '🧪',
    'p2o5': '🟢',
    'elec': '⚡',
    'sim': '📊'
}

def kpi(label, value, unit='', icon_type=None):
    icon_html = f'<div class="kpi-icon">{ICONS[icon_type]}</div>' if icon_type in ICONS else ''
    st.markdown(f'''
        <div class="kpi">
            {icon_html}
            <div class="kpi-content">
                <div class="label">{label}</div>
                <div class="value">{value}</div>
                <div class="unit">{unit}</div>
            </div>
        </div>
    ''', unsafe_allow_html=True)

def fig_layout(fig, title, ytitle='', height=390, xtitle='Temps (h)'):
    fig.update_layout(
        title=dict(text=title, x=.02, font=dict(size=18, color=NAVY, family='Inter')),
        paper_bgcolor='white', plot_bgcolor='white', height=height,
        margin=dict(l=35,r=20,t=58,b=42),
        font=dict(color=TEXT, family='Inter'), hovermode='x unified',
        legend=dict(orientation='h', y=1.05, x=1, xanchor='right', bgcolor='rgba(255,255,255,0)'),
        xaxis=dict(title=xtitle, gridcolor=GRID, zeroline=False, linecolor='#D0D5DD'),
        yaxis=dict(title=ytitle, gridcolor=GRID, zeroline=False, linecolor='#D0D5DD'),
    )
    return fig

def add_phase_bands(fig, df, xcol):
    if 'Phase' not in df.columns: return fig
    phases = df[[xcol,'Phase']].dropna().groupby('Phase', sort=False)[xcol].agg(['min','max']).reset_index()
    cols = ['rgba(10,163,92,.07)','rgba(37,99,235,.06)','rgba(249,115,22,.07)','rgba(234,179,8,.08)','rgba(239,68,68,.06)']
    for i,r in phases.iterrows():
        fig.add_vrect(x0=r['min'], x1=r['max'], fillcolor=cols[i%len(cols)], line_width=0, layer='below')
    return fig

def lines(df, x, cols, title, y='Température (°C)', threshold=False, phases=False, x_range=None, height=390):
    fig = go.Figure()
    palette = [BLUE, RED, ORANGE, GREEN, YELLOW, '#7C3AED', NAVY]
    for i,c in enumerate(cols):
        if c in df.columns:
            fig.add_trace(go.Scatter(x=df[x], y=df[c], mode='lines', name=c.replace('_',' '), line=dict(width=3, shape='spline', color=palette[i%len(palette)])))
    if threshold:
        fig.add_hline(y=T_TARGET, line_dash='dash', line_color='#667085', annotation_text='Objectif 120°C')
    if phases: add_phase_bands(fig, df, x)
    if x_range: fig.update_xaxes(range=x_range)
    return fig_layout(fig, title, y, height)

def area_compare():
    fig = go.Figure()
    if not anc_gas.empty:
        fig.add_trace(go.Scatter(x=anc_gas['Temps_global_h'], y=anc_gas['Conso_gasoil_cumulee_phase_m3'], name='Ancien procédé', line=dict(color=RED, width=3), fill='tozeroy'))
    if not new_gas.empty:
        fig.add_trace(go.Scatter(x=new_gas['Temps_h'], y=new_gas['Conso_gasoil_cumulee_phase_m3'], name='Nouvelle solution', line=dict(color=BLUE, width=3), fill='tozeroy'))
    return fig_layout(fig, 'Consommation cumulée de gasoil', 'm³', 420)

def bar_phase_gasoil():
    if anc_gas.empty or 'Phase' not in anc_gas.columns: return go.Figure()
    a = anc_gas.groupby('Phase', as_index=False)['Debit_gasoil_m3_h'].mean(); a['Procédé']='Ancien'
    n = new_gas.groupby('Phase', as_index=False)['Debit_gasoil_m3_h'].mean(); n['Procédé']='Nouveau'
    d = pd.concat([a,n], ignore_index=True)
    fig = px.bar(d, x='Phase', y='Debit_gasoil_m3_h', color='Procédé', barmode='group', color_discrete_map={'Ancien':RED,'Nouveau':BLUE})
    return fig_layout(fig, 'Débit moyen de gasoil par phase', 'm³/h', 420, '')

def synoptic(t):
    row = new_m.iloc[(new_m['Temps_h'] - t).abs().argmin()]
    T_M4 = row['M4_C']
    T_M3 = row['M3_C']
    T_M2 = row['M2_C']
    T_M1 = row['M1_C']
    active = t <= T_ARRET

    def temp_color(v):
        if v < 50: return "#22C55E"
        if v < 100: return "#EAB308"
        if v < 150: return "#F97316"
        return "#DC2626"

    components.html(f"""
<div style="background:#f8fbff;border:1px solid #dbe5ef;border-radius:18px;padding:8px 14px;">
<div style="font-weight:800;color:#062b4f;font-size:18px;margin-bottom:0px;">
Architecture de la solution vapeur MP / air sec
</div>

<svg viewBox="0 135 1500 855" width="100%" height="855">

<defs>
  <marker id="ab" markerWidth="6" markerHeight="6" refX="5" refY="3" orient="auto">
    <path d="M0,0 L0,6 L6,3 z" fill="#1D4ED8"/>
  </marker>
  <marker id="ar" markerWidth="6" markerHeight="6" refX="5" refY="3" orient="auto">
    <path d="M0,0 L0,6 L6,3 z" fill="#DC2626"/>
  </marker>
  <linearGradient id="metal" x1="0" x2="1">
    <stop offset="0%" stop-color="#cfd6dd"/>
    <stop offset="50%" stop-color="#f4f6f8"/>
    <stop offset="100%" stop-color="#9aa6b2"/>
  </linearGradient>
</defs>

<rect x="25" y="25" width="230" height="90" rx="12" fill="white" stroke="#cbd5e1"/>
<line x1="45" y1="50" x2="100" y2="50" stroke="#1D4ED8" stroke-width="4" marker-end="url(#ab)"/>
<text x="115" y="55" font-size="14" fill="#062b4f" font-weight="700">Air sec</text>
<line x1="45" y1="85" x2="100" y2="85" stroke="#DC2626" stroke-width="4" marker-end="url(#ar)"/>
<text x="115" y="90" font-size="14" fill="#062b4f" font-weight="700">Vapeur MP</text>

<text x="615" y="155" font-size="20" font-weight="800" fill="#062b4f">13C01</text>
<rect x="520" y="180" width="230" height="470" rx="15" fill="url(#metal)" stroke="#334155" stroke-width="3"/>
<polygon points="520,180 635,130 750,180" fill="#eef2f6" stroke="#334155" stroke-width="3"/>

<rect x="545" y="220" width="180" height="65"  fill="{temp_color(T_M1)}" stroke="#263238" stroke-width="2"/>
<rect x="545" y="300" width="180" height="75"  fill="{temp_color(T_M2)}" stroke="#263238" stroke-width="2"/>
<rect x="545" y="390" width="180" height="85"  fill="{temp_color(T_M3)}" stroke="#263238" stroke-width="2"/>
<rect x="545" y="490" width="180" height="115" fill="{temp_color(T_M4)}" stroke="#263238" stroke-width="2"/>

<text x="620" y="260" font-size="22" font-weight="800" fill="white">M1</text>
<text x="620" y="345" font-size="22" font-weight="800" fill="white">M2</text>
<text x="620" y="440" font-size="22" font-weight="800" fill="white">M3</text>
<text x="620" y="555" font-size="22" font-weight="800" fill="white">M4</text>

<text x="765" y="255" font-size="14" font-weight="700" fill="#062b4f">{T_M1:.1f} °C</text>
<text x="765" y="345" font-size="14" font-weight="700" fill="#062b4f">{T_M2:.1f} °C</text>
<text x="765" y="440" font-size="14" font-weight="700" fill="#062b4f">{T_M3:.1f} °C</text>
<text x="765" y="555" font-size="14" font-weight="700" fill="#062b4f">{T_M4:.1f} °C</text>

<text x="565" y="700" font-size="16" font-weight="800" fill="#00a35c">Convertisseur</text>

<text x="165" y="205" font-size="18" font-weight="800" fill="#062b4f">13E01</text>
<rect x="130" y="220" width="110" height="230" rx="30" fill="#ead8b8" stroke="#7c5a2e" stroke-width="2"/>
<text x="135" y="495" font-size="13" font-weight="700" fill="#00a35c">Surchauffeur 1</text>

<text x="355" y="335" font-size="18" font-weight="800" fill="#062b4f">13E02</text>
<rect x="320" y="350" width="110" height="200" rx="30" fill="#ead8b8" stroke="#7c5a2e" stroke-width="2"/>
<text x="315" y="595" font-size="13" font-weight="700" fill="#00a35c">Surchauffeur 2</text>

<text x="1145" y="190" font-size="18" font-weight="800" fill="#062b4f">13E03</text>
<rect x="1100" y="205" width="130" height="270" rx="35" fill="#ead8b8" stroke="#7c5a2e" stroke-width="2"/>
<text x="1110" y="520" font-size="14" font-weight="700" fill="#062b4f">Économiseur</text>

<path d="M1125 240 H1205 V255 H1125 V270 H1205 V285 H1125 V300 H1205 V315 H1125 V330 H1205 V345 H1125" stroke="#DC2626" stroke-width="3.5" fill="none" opacity="0.9" />

<path d="M1125 240 H1070" stroke="#DC2626" stroke-width="4" fill="none" />

<path d="M1125 365 H1205 V377 H1125 V389 H1205 V401 H1125 V413 H1205 V425 H1125 V437 H1205 V449 H1125 V461 H1205 V465 H1125" stroke="#00a35c" stroke-width="3" fill="none" opacity="0.9" />

<path d="M1165 710 V465" stroke="#1D4ED8" stroke-width="6" fill="none" marker-end="url(#ab)"/>
<text x="1050" y="745" font-size="16" font-weight="800" fill="#000000">Air sec issu de la tour de séchage</text>

<path d="M1165 205 V145 H800 V720 H635 V650" stroke="#1D4ED8" stroke-width="4" fill="none" marker-end="url(#ab)"/>

<path d="M520 388 H450 V570 H375 V550" stroke="#1D4ED8" stroke-width="4" fill="none" marker-end="url(#ab)"/>
<path d="M375 350 V310 H490 V378 H520" stroke="#1D4ED8" stroke-width="4" fill="none" marker-end="url(#ab)"/>
<path d="M520 298 H260 V470 H185 V450" stroke="#1D4ED8" stroke-width="4" fill="none" marker-end="url(#ab)"/>
<path d="M185 220 V190 H490 V288 H520" stroke="#1D4ED8" stroke-width="4" fill="none" marker-end="url(#ab)"/>

<path d="M635 130 V60" stroke="#1D4ED8" stroke-width="6" fill="none" marker-end="url(#ab)"/>
<text x="520" y="40" font-size="18" font-weight="800" fill="#000000">Air en sortie de convertisseur</text>

<path d="M1125 790 H1350 V345 H1240" stroke="#DC2626" stroke-width="4" fill="none" />
<text x="1080" y="815" font-size="17" font-weight="800" fill="#000000">Vapeur MP (T=290 °C, P=9 bars)</text>

<line x1="1280" y1="345" x2="1235" y2="345" stroke="#DC2626" stroke-width="4" marker-end="url(#ar)" />
<path d="M1240 345 H1125" stroke="#DC2626" stroke-width="4" fill="none" />

<path d="M1070 240 H1040 V765 H300 V500 H355" stroke="#DC2626" stroke-width="4" fill="none" marker-end="url(#ar)"/>
<path d="M355 500 H405 V475 H355 V450 H405 V425 H355 V400 H405 V375 H355" stroke="#DC2626" stroke-width="3.5" fill="none"/>
<path d="M355 375 H275 V540 H110 V415 H155" stroke="#DC2626" stroke-width="4" fill="none" marker-end="url(#ar)"/>
<path d="M155 415 H205 V390 H155 V365 H205 V340 H155 V315 H205 V290 H155 V265 H205 V245 H155" stroke="#DC2626" stroke-width="3.5" fill="none"/>
<path d="M155 245 H60 V200 H30" stroke="#DC2626" stroke-width="4" fill="none" marker-end="url(#ar)"/>
<text x="30" y="180" font-size="14" font-weight="800" fill="#DC2626">Vers centrale thermique</text>

<rect x="50" y="780" width="220" height="40" rx="8" fill="white" stroke="#cbd5e1"/>
<text x="70" y="805" font-size="14" font-weight="800" fill="#062b4f">Air sec : {'ACTIVE' if active else 'OFF'}</text>

<rect x="290" y="780" width="220" height="40" rx="8" fill="white" stroke="#cbd5e1"/>
<text x="310" y="805" font-size="14" font-weight="800" fill="#062b4f">Vapeur MP : {'ACTIVE' if active else 'OFF'}</text>

<rect x="50" y="835" width="460" height="105" rx="12" fill="white" stroke="#cbd5e1"/>
<text x="70" y="860" font-size="16" font-weight="800" fill="#062b4f">Légende thermique</text>

<circle cx="75" cy="890" r="8" fill="#DC2626"/>
<text x="95" y="895" font-size="14" fill="#062b4f">> 150 °C : très chaud</text>
<circle cx="75" cy="920" r="8" fill="#F97316"/>
<text x="95" y="925" font-size="14" fill="#062b4f">100 – 150 °C : chaud</text>
<circle cx="300" cy="890" r="8" fill="#EAB308"/>
<text x="320" y="895" font-size="14" fill="#062b4f">50 – 100 °C : moyen</text>
<circle cx="300" cy="920" r="8" fill="#22C55E"/>
<text x="320" y="925" font-size="14" fill="#062b4f">< 50 °C : froid</text>

<rect x="1130" y="840" width="280" height="40" rx="8" fill="white" stroke="#cbd5e1"/>
<text x="1150" y="865" font-size="14" font-weight="800" fill="#062b4f">Temps de chauffage : {t:.2f} h</text>

</svg>
</div>
""", height=720, scrolling=False)

# Sidebar
PATH_LOGO = BASE / 'logo_ocp.png'

if PATH_LOGO.exists():
    c1, c2, c3 = st.sidebar.columns([1,2,1])

    with c2:
       st.image(str(PATH_LOGO), width=180)
else:
    st.sidebar.markdown("<h2 style='text-align:center;color:white;'>OCP</h2>", unsafe_allow_html=True)

st.sidebar.markdown(
    """
    <div style="text-align:center; margin-top:6px; margin-bottom:5px;">
        <div style="font-size:15px; font-weight:800; color:white;">OCP SAFI</div>
    </div>
    """,
    unsafe_allow_html=True
)

page_labels = {
    "Vue d'ensemble du procédé": "Vue d'ensemble du procédé",
    "Vapeur & Air": "Préchauffage vapeur-air",
    "Évolution des Masses": "Cinétique des masses",
    "Température du Four": "Profil thermique du four",
    "Analyse énergétique": "Analyse énergétique",
    "Analyse Économique": "Analyse économique"
}

page = st.sidebar.radio(
    "",
    list(page_labels.keys()),
    format_func=lambda x: x,
    label_visibility="collapsed"
)

st.sidebar.markdown("---")
st.sidebar.caption("Dashboard PFE — OCP Safi")

# =====================================================================
# EN-TÊTE PRINCIPAL OPTIMISÉ (PARFAITEMENT VISIBLE)
# =====================================================================
st.markdown(f"""<div style="margin-top:0px;margin-bottom:22px;padding:10px 0px;text-align:center;">
<div style="font-size:28px;font-weight:800;color:{NAVY};margin:0;line-height:1.25;letter-spacing:0.2px;">
Optimisation énergétique et temporelle des lignes sulfuriques
</div>
<div style="margin-top:8px;color:{MUTED};font-size:14px;font-weight:600;">
OCP Safi — Projet de fin d'études
</div>
</div>""", unsafe_allow_html=True)

if page == "Vue d'ensemble du procédé":
    st.markdown(f"""
    <div style="
    background:white;
    border-radius:28px;
    padding:28px;
    margin-bottom:22px;
    border:1px solid #E5E7EB;
    ">

    <h1 style="
    color:{NAVY};
    font-size:30px;
    font-weight:800;
    margin-bottom:10px;
    ">
    Solution optimisée de démarrage à froid
    </h1>

    <p style="
    font-size:17px;
    color:{MUTED};
    line-height:1.7;
    font-weight:500;
    margin-bottom:0px;
    ">
    Chauffage progressif des masses catalytiques par air sec préchauffé à la vapeur moyenne pression.
    </p>

    </div>
    """, unsafe_allow_html=True)

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        kpi("Gain temps", "25 h 21 min", "par démarrage", "time")
    with k2:
        kpi("Gasoil économisé", "39.89", "m³ / démarrage", "gasoil")
    with k3:
        kpi("Gain économique", "579 506,36", "DH / démarrage", "gasoil")
    with k4:
        kpi("Électricité", "357.36", "MWh équivalent", "elec")

    st.markdown("<br>", unsafe_allow_html=True)
    # Remise en place de la version d'origine (simple, autonome et sans clé partagée)
    t = st.slider('Temps de simulation synoptique (h)', 0.0, float(new_m['Temps_h'].max()) if not new_m.empty else 50.0, float(T_ARRET), 0.05)
    synoptic(t)
    # Création de 3 colonnes pour accueillir les structures de données
    col_etat, col_calcul, col_hypothese = st.columns(3)
    
    with col_etat:
        st.markdown(f'<div class="small-title">{ICONS["sim"]} État du système</div>', unsafe_allow_html=True)
        
        # Logique dynamique calquée sur la synoptique (T_ARRET = 11.47 h)
        if t <= T_ARRET:
            st.markdown('<div class="badge-on" style="display:block; text-align:center; margin-bottom:8px;">Air sec : ACTIVE</div>', unsafe_allow_html=True)
            st.markdown('<div class="badge-on" style="display:block; text-align:center; margin-bottom:8px;">Vapeur MP : ACTIVE</div>', unsafe_allow_html=True)
            st.markdown('<div class="badge-on" style="display:block; text-align:center; margin-bottom:8px;">13E03 (Économiseur) : ON</div>', unsafe_allow_html=True)
            st.markdown('<div class="badge-on" style="display:block; text-align:center; margin-bottom:8px;">E02 (Surchauffeur 2) : ON</div>', unsafe_allow_html=True)
            st.markdown('<div class="badge-on" style="display:block; text-align:center; margin-bottom:8px;">E01 (Surchauffeur 1) : ON</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="badge-off" style="display:block; text-align:center; margin-bottom:8px;">Air sec : OFF / ARRÊTÉ</div>', unsafe_allow_html=True)
            st.markdown('<div class="badge-off" style="display:block; text-align:center; margin-bottom:8px;">Vapeur MP : OFF / ARRÊTÉ</div>', unsafe_allow_html=True)
            st.markdown('<div class="badge-off" style="display:block; text-align:center; margin-bottom:8px;">13E03 (Économiseur) : OFF</div>', unsafe_allow_html=True)
            st.markdown('<div class="badge-off" style="display:block; text-align:center; margin-bottom:8px;">E02 (Surchauffeur 2) : OFF</div>', unsafe_allow_html=True)
            st.markdown('<div class="badge-off" style="display:block; text-align:center; margin-bottom:8px;">E01 (Surchauffeur 1) : OFF</div>', unsafe_allow_html=True)

    with col_calcul:
        st.markdown(f'<div class="small-title">{ICONS["sim"]} Données de calcul</div>', unsafe_allow_html=True)
        df_calcul = pd.DataFrame({
            'Paramètre': ['Débit air', 'Débit vapeur', 'Vapeur MP', 'Vitesse turbine'],
            'Valeur': ['3.083 kg/s', '3.854852 kg/s', '290 °C / 9 bar', '4500 tr/min']
        })
        st.dataframe(df_calcul, use_container_width=True, hide_index=True)

    with col_hypothese:
        st.markdown(f'<div class="small-title">{ICONS["sim"]} 🏭 Caractéristiques procédé</div>', unsafe_allow_html=True)
        st.markdown(f"""
        <div style="
         background:{CARD};
         border-radius:18px;
         padding:18px 20px;
         border:1px solid #E5E7EB;
         box-shadow:0 8px 22px rgba(16,24,40,0.06);
         min-height:260px;
        ">

        <div style="
            font-size:14px;
            font-weight:800;
            color:{NAVY};
            margin-bottom:10px;
            text-transform:uppercase;
            letter-spacing:0.4px;
        ">
            Convertisseur cylindrique
        </div>

        <div style="display:flex;justify-content:space-between;margin-bottom:7px;">
            <span style="color:{MUTED};font-weight:600;">Diamètre nominal</span>
            <span style="font-weight:800;color:{TEXT};">14 m</span>
        </div>

        <div style="display:flex;justify-content:space-between;margin-bottom:12px;">
            <span style="color:{MUTED};font-weight:600;">Hauteur nominale</span>
            <span style="font-weight:800;color:{TEXT};">16 m</span>
        </div>

        <div style="
            font-size:14px;
            font-weight:800;
            color:{NAVY};
            margin:14px 0 10px 0;
            text-transform:uppercase;
            letter-spacing:0.4px;
        ">
            Volume de vanadium par masse
        </div>

        <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;">
            <div style="background:#F8FAFC;border-radius:10px;padding:9px 10px;">
                <div style="font-size:12px;color:{MUTED};font-weight:700;">Masse 1</div>
                <div style="font-size:17px;font-weight:800;color:{OCP};">74.1 m³</div>
            </div>
            <div style="background:#F8FAFC;border-radius:10px;padding:9px 10px;">
                <div style="font-size:12px;color:{MUTED};font-weight:700;">Masse 2</div>
                <div style="font-size:17px;font-weight:800;color:{OCP};">80.5 m³</div>
            </div>
            <div style="background:#F8FAFC;border-radius:10px;padding:9px 10px;">
                <div style="font-size:12px;color:{MUTED};font-weight:700;">Masse 3</div>
                <div style="font-size:17px;font-weight:800;color:{OCP};">97.4 m³</div>
            </div>
            <div style="background:#F8FAFC;border-radius:10px;padding:9px 10px;">
                <div style="font-size:12px;color:{MUTED};font-weight:700;">Masse 4</div>
                <div style="font-size:17px;font-weight:800;color:{OCP};">108 m³</div>
            </div>
        </div>

        <div style="
            font-size:14px;
            font-weight:800;
            color:{NAVY};
            margin:16px 0 10px 0;
            text-transform:uppercase;
            letter-spacing:0.4px;
        ">
            Surfaces d'échange thermique
        </div>

        <div style="display:flex;justify-content:space-between;margin-bottom:7px;">
            <span style="color:{MUTED};font-weight:600;">13E03 — serpentin surchauffeur</span>
            <span style="font-weight:800;color:{TEXT};">139 m²</span>
        </div>

        <div style="display:flex;justify-content:space-between;margin-bottom:7px;">
            <span style="color:{MUTED};font-weight:600;">13E02 — surchauffeur 2</span>
            <span style="font-weight:800;color:{TEXT};">142 m²</span>
        </div>

        <div style="display:flex;justify-content:space-between;">
            <span style="color:{MUTED};font-weight:600;">13E01 — surchauffeur 1</span>
            <span style="font-weight:800;color:{TEXT};">1250 m²</span>
        </div>

    </div>
    """, unsafe_allow_html=True)
elif page == "Évolution des Masses":
    st.markdown(f"""
    <div style="
    font-size:22px;
    font-weight:800;
    color:{NAVY};
    margin-bottom:8px;
    ">
    Cinétique thermique des masses catalytiques
    </div>

    <div style="
    font-size:15px;
    color:{MUTED};
    margin-bottom:20px;
    line-height:1.6;
    ">
    Analyse comparative de la montée en température des quatre lits catalytiques durant les différentes phases de démarrage.
    </div>
    """, unsafe_allow_html=True)
    
    # Structure en Onglets du Dashboard
    tab1, tab2, tab3 = st.tabs(['Ancien procédé', 'Nouvelle solution', 'Comparaison'])

    # Configuration des variables de masses et couleurs associées
    cols_masses = ['M1_C', 'M2_C', 'M3_C', 'M4_C']
    noms_masses = {'M1_C': 'Masse 1', 'M2_C': 'Masse 2', 'M3_C': 'Masse 3', 'M4_C': 'Masse 4'}
    couleurs_masses = {'M1_C': GREEN, 'M2_C': BLUE, 'M3_C': ORANGE, 'M4_C': RED}
    y_labels_heat = ['Masse 4', 'Masse 3', 'Masse 2', 'Masse 1']

    # Fonction locale pour abréger les étiquettes de phase
    def abreger_phase(nom_brut):
        nom_clean = str(nom_brut).lower()
        if "chauffage_1" in nom_clean or "chauffage 1" in nom_clean:
            return "Chauffage 1"
        elif "chauffage_2" in nom_clean or "chauffage 2" in nom_clean:
            return "Chauffage 2"
        elif "premiere" in nom_clean or "1ere" in nom_clean or ("manoeuvre" in nom_clean and "1" in nom_clean):
            return "Mnv 1"
        elif "derniere" in nom_clean or "manoeuvre" in nom_clean or "manœuvre" in nom_clean:
            return "Mnv 2"
        else:
            return str(nom_brut).replace("_", " ")

    # ====================================================================
    # TAB 1 : ANCIEN PROCÉDÉ
    # ====================================================================
    with tab1:
        if not anc_m.empty:
            t_max_anc = anc_m['Temps_global_h'].max()
            marge_anc = t_max_anc * 0.05
            range_x_anc = [0, t_max_anc + marge_anc]

            # 1. Graphique Principal
            fig_anc = go.Figure()
            for col_m in cols_masses:
                if col_m in anc_m.columns:
                    fig_anc.add_trace(go.Scatter(
                        x=anc_m['Temps_global_h'],
                        y=anc_m[col_m],
                        name=noms_masses[col_m],
                        mode='lines',
                        line=dict(color=couleurs_masses[col_m], width=2.5)
                    ))
            
            # Ajout des lignes verticales et des textes condensés
            if 'Phase' in anc_m.columns:
                changements_anc = anc_m[anc_m['Phase'] != anc_m['Phase'].shift(-1)]
                for idx, row in changements_anc.iterrows():
                    t_fin = row['Temps_global_h']
                    nom_court = abreger_phase(row['Phase'])
                    
                    if idx != anc_m.index[-1]:
                        fig_anc.add_vline(x=t_fin, line_dash="dash", line_color="#98A2B3", line_width=1.5)
                    
                    df_phase_courante = anc_m[anc_m['Phase'] == row['Phase']]
                    t_debut = df_phase_courante['Temps_global_h'].min()
                    t_centre = t_debut + (df_phase_courante['Temps_global_h'].max() - t_debut) / 2
                    
                    fig_anc.add_annotation(
                        x=t_centre, y=1.02, yref='paper',
                        text=f"⬇ {nom_court}", showarrow=False,
                        font=dict(size=10, color=NAVY, weight="bold"),
                        xanchor='center', yanchor='bottom'
                    )

            fig_anc.update_layout(
                title="Évolution en températures — Ancien procédé",
                xaxis=dict(title="Temps global (h)", range=range_x_anc),
                yaxis_title="Température (°C)",
                hovermode="x unified",
                margin=dict(t=90), 
                legend=dict(orientation="h", yanchor="top", y=0.98, yref="container", xanchor="center", x=0.5)
            )
            st.plotly_chart(fig_anc, width='stretch')
            
            st.markdown("---")
            
            # 2. Carte thermique
            df_sample_anc = anc_m.iloc[::10] if len(anc_m) > 200 else anc_m
            z_anc = [
                df_sample_anc['M4_C'].values if 'M4_C' in df_sample_anc.columns else np.zeros(len(df_sample_anc)),
                df_sample_anc['M3_C'].values if 'M3_C' in df_sample_anc.columns else np.zeros(len(df_sample_anc)),
                df_sample_anc['M2_C'].values if 'M2_C' in df_sample_anc.columns else np.zeros(len(df_sample_anc)),
                df_sample_anc['M1_C'].values if 'M1_C' in df_sample_anc.columns else np.zeros(len(df_sample_anc))
            ]
            fig_heat_anc = go.Figure(data=go.Heatmap(
                z=z_anc, x=df_sample_anc['Temps_global_h'], y=y_labels_heat,
                colorscale='Thermal', colorbar=dict(title="°C")
            ))
            fig_heat_anc.update_layout(
                title="Carte thermique spatio-temporelle — Ancien procédé",
                xaxis=dict(title="Temps global (h)", range=range_x_anc),
                yaxis_title="Position des Lits",
                height=260
            )
            st.plotly_chart(fig_heat_anc, width='stretch')

            # Section Durée Réelle des Phases
            st.markdown(f'<div class="small-title">{ICONS["time"]} Chronologie et Durée Réelle des Phases (Ancien)</div>', unsafe_allow_html=True)
            if 'Phase' in anc_m.columns:
                phases_anc = anc_m['Phase'].unique()
                cols_cards = st.columns(len(phases_anc) if len(phases_anc) > 0 else 1)
                for idx, p_nom in enumerate(phases_anc):
                    df_p = anc_m[anc_m['Phase'] == p_nom]
                    duree = df_p['Temps_global_h'].max() - df_p['Temps_global_h'].min()
                    nom_propre = p_nom.replace("_", " ")
                    if "manoeuvre" in str(p_nom).lower() or "manœuvre" in str(p_nom).lower():
                        h, m = 5, 0
                    else:
                        h = int(duree)
                        m = round((duree - h) * 60)
                    with cols_cards[idx]:
                        st.markdown(f'''
                        <div style="background-color: {CARD}; border-left: 5px solid {NAVY}; padding: 15px; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); height: 125px;">
                            <div style="font-size: 11px; text-transform: uppercase; color: {MUTED}; font-weight: 600; min-height: 34px; line-height: 1.2;">{nom_propre}</div>
                            <div style="font-size: 20px; font-weight: 700; color: {TEXT}; margin-top: 5px;">{h}h {m:02d}min</div>
                            <div style="font-size: 11px; color: {MUTED}; margin-top: 2px;">De {df_p["Temps_global_h"].min():.1f}h à {df_p["Temps_global_h"].max():.1f}h</div>
                        </div>
                        ''', unsafe_allow_html=True)
        else:
            st.warning("Fichier de données de l'ancien procédé indisponible.")

    # ====================================================================
    # TAB 2 : NOUVELLE SOLUTION
    # ====================================================================
    with tab2:
        if not new_m.empty:
            t_max_new = new_m['Temps_h'].max()
            marge_new = t_max_new * 0.05
            range_x_new = [0, t_max_new + marge_new]

            # 1. Graphique Principal
            fig_new = go.Figure()
            for col_m in cols_masses:
                if col_m in new_m.columns:
                    fig_new.add_trace(go.Scatter(
                        x=new_m['Temps_h'],
                        y=new_m[col_m],
                        name=noms_masses[col_m],
                        mode='lines',
                        line=dict(color=couleurs_masses[col_m], width=2.5)
                    ))
            
            # Ajout des lignes verticales et des textes condensés
            if 'Phase' in new_m.columns:
                changements_new = new_m[new_m['Phase'] != new_m['Phase'].shift(-1)]
                for idx, row in changements_new.iterrows():
                    t_fin = row['Temps_h']
                    nom_court = abreger_phase(row['Phase'])
                    
                    if idx != new_m.index[-1]:
                        fig_new.add_vline(x=t_fin, line_dash="dash", line_color="#98A2B3", line_width=1.5)
                    
                    df_phase_courante = new_m[new_m['Phase'] == row['Phase']]
                    t_debut = df_phase_courante['Temps_h'].min()
                    t_centre = t_debut + (df_phase_courante['Temps_h'].max() - t_debut) / 2
                    
                    fig_new.add_annotation(
                        x=t_centre, y=1.02, yref='paper',
                        text=f"⬇ {nom_court}", showarrow=False,
                        font=dict(size=10, color=OCP, weight="bold"),
                        xanchor='center', yanchor='bottom'
                    )

            fig_new.update_layout(
                title="Évolution en températures — Nouvelle solution",
                xaxis=dict(title="Temps (h)", range=range_x_new),
                yaxis_title="Température (°C)",
                hovermode="x unified",
                margin=dict(t=90),
                legend=dict(orientation="h", yanchor="top", y=0.98, yref="container", xanchor="center", x=0.5)
            )
            st.plotly_chart(fig_new, width='stretch')
            
            st.markdown("---")
            
            # 2. Carte thermique
            df_sample_new = new_m.iloc[::10] if len(new_m) > 200 else new_m
            z_new = [
                df_sample_new['M4_C'].values if 'M4_C' in df_sample_new.columns else np.zeros(len(df_sample_new)),
                df_sample_new['M3_C'].values if 'M3_C' in df_sample_new.columns else np.zeros(len(df_sample_new)),
                df_sample_new['M2_C'].values if 'M2_C' in df_sample_new.columns else np.zeros(len(df_sample_new)),
                df_sample_new['M1_C'].values if 'M1_C' in df_sample_new.columns else np.zeros(len(df_sample_new))
            ]
            fig_heat_new = go.Figure(data=go.Heatmap(
                z=z_new, x=df_sample_new['Temps_h'], y=y_labels_heat,
                colorscale='Thermal', colorbar=dict(title="°C")
            ))
            fig_heat_new.update_layout(
                title="Carte thermique spatio-temporelle — Nouvelle solution",
                xaxis=dict(title="Temps (h)", range=range_x_new),
                yaxis_title="Position des Lits",
                height=260
            )
            st.plotly_chart(fig_heat_new, width='stretch')

            # Section Durée Réelle des Phases
            st.markdown(f'<div class="small-title">{ICONS["time"]} Chronologie et Durée Réelle des Phases (Nouveau)</div>', unsafe_allow_html=True)
            if 'Phase' in new_m.columns:
                phases_new = new_m['Phase'].unique()
                cols_cards = st.columns(len(phases_new) if len(phases_new) > 0 else 1)
                for idx, p_nom in enumerate(phases_new):
                    df_p = new_m[new_m['Phase'] == p_nom]
                    duree = df_p['Temps_h'].max() - df_p['Temps_h'].min()
                    
                    nom_propre = p_nom.replace("_", " ")
                    if "manoeuvre" in str(p_nom).lower() or "manœuvre" in str(p_nom).lower():
                        h, m = 5, 0
                    else:
                        h = int(duree)
                        m = round((duree - h) * 60)
                    with cols_cards[idx]:
                        st.markdown(f'''
                        <div style="background-color: {CARD}; border-left: 5px solid {OCP}; padding: 15px; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); height: 125px;">
                            <div style="font-size: 11px; text-transform: uppercase; color: {MUTED}; font-weight: 600; min-height: 34px; line-height: 1.2;">{nom_propre}</div>
                            <div style="font-size: 20px; font-weight: 700; color: {TEXT}; margin-top: 5px;">{h}h {m:02d}min</div>
                            <div style="font-size: 11px; color: {MUTED}; margin-top: 2px;">De {df_p["Temps_h"].min():.1f}h à {df_p["Temps_h"].max():.1f}h</div>
                        </div>
                        ''', unsafe_allow_html=True)

    # ====================================================================
    # TAB 3 : COMPARAISON DIRECTE DU LIT CRITIQUE (M3)
    # ====================================================================
    with tab3:
        t_max_comp = max(
            anc_m['Temps_global_h'].max() if 'Temps_global_h' in anc_m.columns else 0,
            new_m['Temps_h'].max() if 'Temps_h' in new_m.columns else 0
        )
        range_x_comp = [0, t_max_comp + (t_max_comp * 0.05)]

        fig_comp = go.Figure()
        
        if 'Temps_global_h' in anc_m.columns and 'M3_C' in anc_m.columns:
            fig_comp.add_trace(go.Scatter(
                x=anc_m['Temps_global_h'], y=anc_m['M3_C'],
                name='Ancien M3 (Référence)', mode='lines',
                line=dict(color=RED, width=3)
            ))
            
        if 'Temps_h' in new_m.columns and 'M3_C' in new_m.columns:
            fig_comp.add_trace(go.Scatter(
                x=new_m['Temps_h'], y=new_m['M3_C'],
                name='Nouveau M3 (Optimal)', mode='lines',
                line=dict(color=BLUE, width=3)
            ))
        
        # CORRECTION : Remplacement du libellé pour correspondre à l'objectif réel du procédé
        fig_comp.add_hline(
            y=120, line_dash='dash', line_color='#667085',
            annotation_text="Température objectif de préchauffage (120°C)",
            annotation_position="top left"
        )
        
        fig_comp.update_layout(
            title='Superposition dynamique de la masse limitante M3 (Ancien vs Nouveau)',
            xaxis=dict(title="Temps (h)", range=range_x_comp),
            yaxis_title="Température de la Masse 3 (°C)",
            hovermode="x unified",
            margin=dict(t=70),
            legend=dict(orientation="h", yanchor="top", y=0.98, yref="container", xanchor="center", x=0.5)
        )
        st.plotly_chart(fig_comp, width='stretch')

elif page == 'Température du Four':
    st.markdown(f"""
    <div style="
    font-size:22px;
    font-weight:800;
    color:{NAVY};
    margin-bottom:8px;
    ">
    Suivi dynamique des opérations de démarrage
    </div>
    """, unsafe_allow_html=True)

    # ====================================================================
    # PARTIE 1 : CONFIGURATION DE LA SIMULATION (ANCIEN VS NOUVEAU)
    # ====================================================================
    st.markdown(f'<div class="small-title">🕒 Simulateur Cinématique et Thermique du Procédé</div>', unsafe_allow_html=True)
    
    c_cfg1, c_cfg2 = st.columns([2, 3])
    with c_cfg1:
        type_proc = st.radio("Sélectionner le procédé à simuler :", ["Ancien procédé", "Nouvelle solution"], horizontal=True)
    
    if type_proc == "Ancien procédé":
        df_f_sim = anc_four
        df_g_sim = anc_gas
        col_t = 'Temps_global_h'
    else:
        df_f_sim = new_four
        df_g_sim = new_gas
        col_t = 'Temps_h'

    with c_cfg2:
        max_t_slider = float(df_f_sim[col_t].max()) if not df_f_sim.empty else 10.0
        t_four = st.slider('Ajuster le temps de fonctionnement (h)', 0.0, max_t_slider, 0.0, 0.1, key='slider_four_dynamic')

    # Extraction des lignes de données courantes à l'instant t
    idx_f = (df_f_sim[col_t] - t_four).abs().argmin() if not df_f_sim.empty else 0
    row_f = df_f_sim.iloc[idx_f] if not df_f_sim.empty else {'Temperature_four_C': 20.0, 'Phase': 'Inconnue'}
    
    temp_actuelle = row_f['Temperature_four_C']
    phase_brute = str(row_f['Phase']).lower()

    # ====================================================================
    # GESTION DU GASOIL CUMULÉ
    # ====================================================================
    conso_cumulee = 0.0

    phase_norm = (
        phase_brute
        .replace("œ", "oe")
        .replace("é", "e")
        .replace("è", "e")
        .replace("ê", "e")
        .replace("à", "a")
    )

    if not df_g_sim.empty:
        is_phase_sans_gasoil = (
            "soufflage" in phase_norm
            or "manoeuvre" in phase_norm
            or "derniere" in phase_norm
            or "fin" in phase_norm
        )

        is_chauffage_2 = (
            "chauffage_2" in phase_norm
            or "chauffage 2" in phase_norm
            or "chauffage2" in phase_norm
        )

        if is_phase_sans_gasoil:
            conso_cumulee = 0.00

        elif is_chauffage_2:
            val_max_c1 = 9.96 if type_proc == "Nouvelle solution" else 38.80

            row_g_actuelle = df_g_sim.iloc[(df_g_sim[col_t] - t_four).abs().argmin()]
            conso_cumulee = val_max_c1 + row_g_actuelle["Conso_gasoil_cumulee_phase_m3"]

        else:
            row_g_actuelle = df_g_sim.iloc[(df_g_sim[col_t] - t_four).abs().argmin()]
            conso_cumulee = row_g_actuelle["Conso_gasoil_cumulee_phase_m3"]

    # ====================================================================
    # LOGIQUE OPÉRATOIRE DES FLUIDES & COULEURS DU SYNOPTIQUE
    # ====================================================================
    etat_registre = "DÉMONTÉ (Voie Libre)"
    etat_cheminee = "FERMÉE"
    commentaire_action = ""

    show_smoke_cheminee = False
    show_smoke_convertisseur = False
    is_combustion_active = True

    color_registre = "rgba(0,0,0,0)"
    stroke_registre = "rgba(0,0,0,0)"

    if "manoeuvre" in phase_norm:
        if "derniere" in phase_norm or "fin" in phase_norm:
            etat_registre = "DÉMONTÉ (Voie Libre)"
            etat_cheminee = "FERMÉE"
            commentaire_action = "🛑 Brûleurs arrêtés : Stabilisation thermique finale et étanchéité avant admission du gaz de process."
            show_smoke_cheminee = False
            show_smoke_convertisseur = False
            is_combustion_active = False
        else:
            etat_registre = "EN COURS DE DÉMONTAGE"
            etat_cheminee = "EN COURS DE FERMETURE"
            commentaire_action = "🛑 Brûleurs arrêtés (Pas de fumée) : Pose du joint plein sur la cheminée et retrait du registre d'isolement convertisseur."
            show_smoke_cheminee = False
            show_smoke_convertisseur = False
            is_combustion_active = False
            color_registre = "#eab308"
            stroke_registre = "#a16207"

    elif (
        "chauffage_1" in phase_norm
        or "chauffage 1" in phase_norm
        or "vapeur" in phase_norm
        or "mp" in phase_norm
    ):
        etat_registre = "MONTÉ (Ligne Bloquée)"
        etat_cheminee = "OUVERTE"
        commentaire_action = "🔥 Brûleurs actifs : Évacuation des fumées humides de gasoil par la cheminée ouverte pour protéger le catalyseur."
        show_smoke_cheminee = t_four > 0.01
        show_smoke_convertisseur = False
        is_combustion_active = True
        color_registre = "#ef4444"
        stroke_registre = "#991b1b"

    elif "soufflage" in phase_norm:
        etat_registre = "DÉMONTÉ (Voie Libre)"
        etat_cheminee = "FERMÉE"
        commentaire_action = "💨 Soufflage d'air pur à travers le circuit pour balayer et préchauffer les lits du convertisseur."
        show_smoke_cheminee = False
        show_smoke_convertisseur = True
        is_combustion_active = False

    elif (
        "chauffage_2" in phase_norm
        or "chauffage 2" in phase_norm
        or "chauffage2" in phase_norm
    ):
        etat_registre = "DÉMONTÉ (Voie Libre)"
        etat_cheminee = "FERMÉE"
        commentaire_action = "🔥 Brûleurs réactivés : Les fumées chaudes sont envoyées directement dans le convertisseur."
        show_smoke_cheminee = False
        show_smoke_convertisseur = True
        is_combustion_active = True

    color_cercle_cheminee = "#22c55e" if etat_cheminee == "OUVERTE" else "#ef4444"
    color_texte_cheminee = "#16a34a" if etat_cheminee == "OUVERTE" else "#dc2626"

    def get_fire_color(v, active):
        if not active: return "#475569"
        if v < 250: return "#f97316"
        return "#dc2626"

    # Affichage des KPIs supérieurs
    c_kpi1, c_kpi2, c_kpi3 = st.columns(3)
    with c_kpi1: kpi('Température Four', f"{temp_actuelle:.1f}", "°C", "sim")
    with c_kpi2: kpi('Gasoil Total Cumulé', f"{conso_cumulee:.2f}", "m³", "gasoil")
    with c_kpi3: kpi('Phase en cours', phase_brute.replace("_", " ").upper(), f"Temps: {t_four:.1f}h", "time")

    st.info(f"💡 **Statut Mécanique & Thermique :** {commentaire_action}")

    # ====================================================================
    # SYNOPTIQUE DYNAMIQUE (CONSERVÉ À L'IDENTIQUE)
    # ====================================================================
    components.html(f"""
    <div style="background:#f8fbff; border:1px solid #dbe5ef; border-radius:18px; padding:15px; text-align:center;">
        <svg viewBox="0 0 1050 420" width="100%" height="380">
            <defs>
                <marker id="arrow-blue" markerWidth="6" markerHeight="6" refX="0" refY="2" orient="auto">
                  <path d="M0,0 L0,4 L5,2 z" fill="#2563eb" />
                </marker>
            </defs>

            <rect x="20" y="160" width="160" height="130" rx="15" fill="#334155" stroke="#1e293b" stroke-width="3"/>
            <rect x="32" y="172" width="136" height="106" rx="8" fill="{get_fire_color(temp_actuelle, is_combustion_active)}" opacity="0.85">
                {"<animate attributeName='opacity' values='0.7;0.9;0.7' dur='1.5s' repeatCount='indefinite' />" if is_combustion_active else ""}
            </rect>
            <text x="100" y="232" text-anchor="middle" font-family="Inter" font-weight="bold" fill="white" font-size="22">{temp_actuelle:.1f}°C</text>
            <text x="100" y="315" text-anchor="middle" font-family="Inter" font-size="12" font-weight="bold" fill="#334155">FOUR CYLINDRIQUE</text>

            <path d="M 180 225 L 240 225" stroke="#475569" stroke-width="8" fill="none" />

            <rect x="240" y="140" width="130" height="160" rx="6" fill="#64748b" stroke="#334155" stroke-width="2"/>
            <line x1="265" y1="160" x2="265" y2="280" stroke="#e2e8f0" stroke-width="3" stroke-dasharray="5,5"/>
            <line x1="305" y1="160" x2="305" y2="280" stroke="#e2e8f0" stroke-width="3" stroke-dasharray="5,5"/>
            <line x1="345" y1="160" x2="345" y2="280" stroke="#e2e8f0" stroke-width="3" stroke-dasharray="5,5"/>
            <text x="305" y="315" text-anchor="middle" font-family="Inter" font-size="12" font-weight="bold" fill="#334155">CHAUDIÈRE</text>

            <path d="M 370 225 L 430 225" stroke="#475569" stroke-width="8" fill="none" />

            <polygon points="430,120 540,150 540,270 430,300" fill="#94a3b8" stroke="#334155" stroke-width="2.5" />
            <line x1="485" y1="140" x2="485" y2="280" stroke="#475569" stroke-width="2"/>
            <text x="485" y="325" text-anchor="middle" font-family="Inter" font-size="12" font-weight="bold" fill="#334155">FILTRE À GAZ CHAUD</text>

            <path d="M 540 210 L 620 210" stroke="#475569" stroke-width="8" fill="none" />

            <path d="M 620 210 L 620 80" stroke="#475569" stroke-width="8" fill="none" />
            <rect x="600" y="30" width="40" height="50" fill="#475569" rx="2" />
            
            <circle cx="620" cy="80" r="10" fill="{color_cercle_cheminee}" stroke="white" stroke-width="1.5"/>
            <text x="640" y="45" font-family="Inter" font-size="11" font-weight="bold" fill="{color_texte_cheminee}">
                {f"CHEMINÉE: {etat_cheminee}"}
            </text>
            
            {f'''
            <circle cx="620" cy="20" r="6" fill="#94a3b8" opacity="0.6"><animate attributeName="cy" from="20" to="-20" dur="1.2s" repeatCount="indefinite"/><animate attributeName="opacity" from="0.6" to="0" dur="1.2s" repeatCount="indefinite"/><animate attributeName="r" from="6" to="14" dur="1.2s" repeatCount="indefinite"/></circle>
            <circle cx="630" cy="15" r="4" fill="#cbd5e1" opacity="0.4"><animate attributeName="cy" from="15" to="-25" dur="0.9s" repeatCount="indefinite"/><animate attributeName="opacity" from="0.4" to="0" dur="0.9s" repeatCount="indefinite"/></circle>
            ''' if show_smoke_cheminee else ""}

            <path d="M 620 210 L 780 210" stroke="{"#cbd5e1" if ("chauffage_1" in phase_brute or "vapeur" in phase_brute) else "#2563eb"}" stroke-width="8" fill="none" stroke-dasharray="{"none" if ("chauffage_1" in phase_brute or "vapeur" in phase_brute) else "8,4"}" />
            
            <rect x="690" y="190" width="12" height="40" fill="{color_registre}" stroke="{stroke_registre}" stroke-width="1.5" rx="1"/>
            <text x="696" y="175" text-anchor="middle" font-family="Inter" font-size="10" font-weight="bold" fill="#475569">{f"REGISTRE: {etat_registre}"}</text>

            <rect x="780" y="80" width="150" height="260" rx="20" fill="#052B4F" stroke="#03192e" stroke-width="3"/>
            
            <rect x="795" y="110" width="120" height="20" fill="#0AA35C" rx="2" opacity="0.85"/>
            <text x="855" y="124" text-anchor="middle" font-family="Inter" font-size="10" font-weight="bold" fill="white">Masse 1</text>
            
            <rect x="795" y="160" width="120" height="20" fill="#2563EB" rx="2" opacity="0.85"/>
            <text x="855" y="174" text-anchor="middle" font-family="Inter" font-size="10" font-weight="bold" fill="white">Masse 2</text>
            
            <rect x="795" y="210" width="120" height="20" fill="#F97316" rx="2" opacity="0.85"/>
            <text x="855" y="224" text-anchor="middle" font-family="Inter" font-size="10" font-weight="bold" fill="white">Masse 3</text>
            
            <rect x="795" y="260" width="120" height="20" fill="#EF4444" rx="2" opacity="0.85"/>
            <text x="855" y="274" text-anchor="middle" font-family="Inter" font-size="10" font-weight="bold" fill="white">Masse 4</text>
            
            <text x="855" y="365" text-anchor="middle" font-family="Inter" font-size="12" font-weight="bold" fill="#052B4F">CONVERTISSEUR</text>

            {'''
            <path d="M 750 210 L 780 210" stroke="#2563eb" stroke-width="4" marker-end="url(#arrow-blue)"/>
            <circle cx="810" cy="140" r="3" fill="#38bdf8"><animate attributeName="cy" from="130" to="310" dur="2s" repeatCount="indefinite"/></circle>
            <circle cx="895" cy="140" r="3" fill="#38bdf8"><animate attributeName="cy" from="130" to="310" dur="1.7s" repeatCount="indefinite"/></circle>
            ''' if show_smoke_convertisseur else ""}

        </svg>
    </div>
    """, height=410)

    st.markdown("---")

    # ====================================================================
    # PARTIE 2 : GRAPHIQUE THERMIQUE (CONSERVÉ À L'IDENTIQUE)
    # ====================================================================
    st.markdown(f'<div class="small-title">📈 Profil Thermique Réel et Transitions des Phases</div>', unsafe_allow_html=True)
    
    fig_f = go.Figure()
    
    if type_proc == "Ancien procédé":
        if 'Temps_global_h' in anc_four.columns:
            fig_f.add_trace(go.Scatter(
                x=anc_four['Temps_global_h'], 
                y=anc_four['Temperature_four_C'], 
                name="Four (°C)", 
                line=dict(color='#FF4B4B', width=3)
            ))
            
            # --- ÉTIQUETTES DES PHASES PARFAITEMENT CENTRÉES ---
            fig_f.add_annotation(x=16.16, y=1.05, yref="paper", text="Chauffage 1", showarrow=False, font=dict(size=11, color="#94A3B8", family="Inter"))
            fig_f.add_annotation(x=34.83, y=1.05, yref="paper", text="Mnv", showarrow=False, font=dict(size=11, color="#94A3B8", family="Inter"))
            fig_f.add_annotation(x=39.75, y=1.05, yref="paper", text="Soufflage", showarrow=False, font=dict(size=11, color="#94A3B8", family="Inter"))
            fig_f.add_annotation(x=56.00, y=1.05, yref="paper", text="Chauffage 2", showarrow=False, font=dict(size=11, color="#94A3B8", family="Inter"))
            fig_f.add_annotation(x=72.25, y=1.05, yref="paper", text="Mnv F.", showarrow=False, font=dict(size=11, color="#94A3B8", family="Inter"))
            
    else:
        if 'Temps_h' in new_four.columns:
            fig_f.add_trace(go.Scatter(
                x=new_four['Temps_h'], 
                y=new_four['Temperature_four_C'], 
                name="Four (°C)", 
                line=dict(color='#00E676', width=3)
            ))
            
            # --- ÉTIQUETTES NOUVELLE SOLUTION ---
            fig_f.add_annotation(x=5.66, y=1.05, yref="paper", text="Vapeur MP", showarrow=False, font=dict(size=11, color="#94A3B8", family="Inter"))
            fig_f.add_annotation(x=12.41, y=1.05, yref="paper", text="Mnv", showarrow=False, font=dict(size=11, color="#94A3B8", family="Inter"))
            fig_f.add_annotation(x=28.80, y=1.05, yref="paper", text="Chauffage 2", showarrow=False, font=dict(size=11, color="#94A3B8", family="Inter"))
            fig_f.add_annotation(x=44.30, y=1.05, yref="paper", text="Mnv Fin", showarrow=False, font=dict(size=11, color="#94A3B8", family="Inter"))

    # Curseur temporel de suivi interactif
    fig_f.add_vline(x=t_four, line_width=1.5, line_color="#FFD700", line_dash="solid")

    fig_f.update_layout(
        plot_bgcolor='white',   
        paper_bgcolor='white',  
        font=dict(color='#94A3B8', family="Inter"),
        xaxis=dict(
            tickmode='linear', tick0=0, dtick=10,
            gridcolor='rgba(148, 163, 184, 0.08)', showgrid=True, zeroline=False, autorange=True,
            title=dict(text="Temps global (h)", font=dict(size=12, color="#94A3B8"))
        ),
        yaxis=dict(
            gridcolor='rgba(148, 163, 184, 0.08)', showgrid=True, zeroline=False, autorange=True,
            title=dict(text="Température du Four (°C)", font=dict(size=12, color="#94A3B8"))
        ),
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=-0.22, xanchor="center", x=0.5, font=dict(size=11, color="#94A3B8")),
        margin=dict(t=50, b=60, l=50, r=30)
    )

    st.plotly_chart(fig_f, use_container_width=True)

elif page == 'Vapeur & Air':
    # 1. Rangée des indicateurs clés (KPIs)
    c1, c2, c3, c4 = st.columns(4)
    with c1: kpi('Débit air sec', '3.083', 'kg/s', 'sim')
    with c2: kpi('Débit vapeur MP', '3.854852', 'kg/s', 'sim')
    with c3: kpi('Arrêt des flux', '11 h 28 min', 'Objectif atteint', 'time')
    with c4: kpi('Puissance Éco (13E03)', f"{opt['surchauffeur']['Q_kW']:.2f}", 'kW', 'elec')
    
    st.write('')
    
    # 2. Affichage des deux graphiques côte à côte
    if not new_va.empty and 'Temps_vapeur_air_h' in new_va.columns:
        va = new_va[new_va['Temps_vapeur_air_h'] <= T_ARRET].copy()
        col_graphes_a, col_graphes_b = st.columns(2)
        
        with col_graphes_a:
            st.plotly_chart(lines(va, 'Temps_vapeur_air_h', ['T_air_out_13E03_C', 'T_air_out_E02_C', 'T_air_out_E01_C'], 'Température air dans les échangeurs', '°C', x_range=[0, T_ARRET], height=400), use_container_width=True)
            
        with col_graphes_b:
            st.plotly_chart(lines(va, 'Temps_vapeur_air_h', ['T_vap_out_13E03_C', 'T_vap_out_E02_C', 'T_vap_out_E01_C', 'T_sat_vapeur_C'], 'Température vapeur dans les échangeurs', '°C', x_range=[0, T_ARRET], height=400), use_container_width=True)
            
    st.write('')
    
    # 3. Nouveau Tableau dynamique des performances des échangeurs
    st.markdown(f'<div class="small-title">{ICONS["sim"]} Performances thermiques des Échangeurs</div>', unsafe_allow_html=True)
    
    # Slider indépendant pour la page Vapeur & Air limité à T_ARRET (11.47h)
    t_va = st.slider(
        'Temps de circulation Vapeur & Air Sec (h)', 
        0.0, 
        float(T_ARRET), 
        0.0, 
        0.05, 
        key='slider_independant_vapeur_air'
    )
    
    # Extraction croisée des données
    if not new_va.empty and not new_m.empty:
        # Trouver la ligne correspondante dans la feuille 04_Vapeur_Air
        idx_va = (new_va['Temps_vapeur_air_h'] - t_va).abs().idxmin()
        row_va = new_va.loc[idx_va]
        
        # Trouver la ligne correspondante au même instant t dans la feuille globale/masses (02_Masses)
        idx_m = (new_m['Temps_h'] - t_va).abs().idxmin()
        row_m = new_m.loc[idx_m]
        
        st.write(f"**Analyse instantanée de l'échange thermique à t = {t_va:.2f} h :**")
        
        # ====================================================================
        # LOGIQUE THERMODYNAMIQUE : COUPLAGE ÉCHANGEURS & MASSES (Efficacité 100%)
        # ====================================================================
        
        # --- 13E03 (Économiseur) ---
        t_air_in_13E03 = 53.0
        t_air_out_13E03 = row_va['T_air_out_13E03_C']
        gain_air_13E03 = t_air_out_13E03 - t_air_in_13E03
        
        # --- E02 (Surchauffeur 2) ---
        # L'air entrant provient de la Masse 3 (colonne 'M3_C' de la feuille masses)
        t_air_in_E02 = row_m['M3_C']
        t_air_out_E02 = row_va['T_air_out_E02_C']
        gain_air_E02 = t_air_out_E02 - t_air_in_E02
        
        # --- E01 (Surchauffeur 1) ---
        # L'air entrant provient de la Masse 2 (colonne 'M2_C' de la feuille masses)
        t_air_in_E01 = row_m['M2_C']
        t_air_out_E01 = row_va['T_air_out_E01_C']
        gain_air_E01 = t_air_out_E01 - t_air_in_E01
        
        # ====================================================================
        # ÉVOLUTION THERMIQUE CÔTÉ VAPEUR (Série Vapeur MP)
        # ====================================================================
        t_vap_in_13E03 = 290.0
        t_vap_out_13E03 = row_va['T_vap_out_13E03_C']
        chute_vap_13E03 = t_vap_out_13E03 - t_vap_in_13E03
        
        t_vap_in_E02 = t_vap_out_13E03
        t_vap_out_E02 = row_va['T_vap_out_E02_C']
        chute_vap_E02 = t_vap_out_E02 - t_vap_in_E02
        
        t_vap_in_E01 = t_vap_out_E02
        t_vap_out_E01 = row_va['T_vap_out_E01_C']
        chute_vap_E01 = t_vap_out_E01 - t_vap_in_E01

        # Structuration du tableau dynamique final
        df_perf_dynamique = pd.DataFrame({
            'Flux / Échangeur': [
                '13E03 (Économiseur) - Côté Air', 
                'E02 (Surchauffeur 2) - Côté Air', 
                'E01 (Surchauffeur 1) - Côté Air',
                '13E03 (Économiseur) - Côté Vapeur', 
                'E02 (Surchauffeur 2) - Côté Vapeur', 
                'E01 (Surchauffeur 1) - Côté Vapeur'
            ],
            'Température Entrée (°C)': [
                f"{t_air_in_13E03:.1f}", f"{t_air_in_E02:.1f}", f"{t_air_in_E01:.1f}",
                f"{t_vap_in_13E03:.1f}", f"{t_vap_in_E02:.1f}", f"{t_vap_in_E01:.1f}"
            ],
            'Température Sortie (°C)': [
                f"{t_air_out_13E03:.1f}", f"{t_air_out_E02:.1f}", f"{t_air_out_E01:.1f}",
                f"{t_vap_out_13E03:.1f}", f"{t_vap_out_E02:.1f}", f"{t_vap_out_E01:.1f}"
            ],
            'Évolution Thermique (ΔT)': [
                f"+{gain_air_13E03:.1f} °C", f"+{gain_air_E02:.1f} °C", f"+{gain_air_E01:.1f} °C",
                f"{chute_vap_13E03:.1f} °C", f"{chute_vap_E02:.1f} °C", f"{chute_vap_E01:.1f} °C"
            ]
        })
        
        st.dataframe(df_perf_dynamique, use_container_width=True, hide_index=True)
        st.caption(f"💡 Vérification : À t = {t_va:.2f}h, T_entrée E02 = Masse 3 ({t_air_in_E02:.1f}°C) | T_entrée E01 = Masse 2 ({t_air_in_E01:.1f}°C).")
    else:
        st.warning("Données indisponibles dans new_m ou new_va.")

elif page == 'Analyse énergétique':

    st.markdown(f"""
    <div style="font-size:22px;font-weight:800;color:{NAVY};margin-bottom:8px;">
    Analyse énergétique du combustible
    </div>
    <div style="font-size:15px;color:{MUTED};margin-bottom:20px;line-height:1.6;">
    Évaluation comparative de la consommation de gasoil entre l’ancien procédé et la nouvelle solution de démarrage.
    </div>
    """, unsafe_allow_html=True)

    ancien_ch1 = 38.80
    ancien_ch2 = 31.20
    ancien_total = 70.00

    nouveau_ch1 = 9.96
    nouveau_ch2 = 20.15
    nouveau_total = 30.11

    economie = ancien_total - nouveau_total
    reduction = round(economie / ancien_total * 100, 2)

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        kpi("Ancien procédé", f"{ancien_total:.2f}", "m³", "gasoil")
    with c2:
        kpi("Nouvelle solution", f"{nouveau_total:.2f}", "m³", "gasoil")
    with c3:
        kpi("Gasoil économisé", f"{economie:.2f}", "m³ / démarrage", "gasoil")
    with c4:
        kpi("Réduction obtenue", f"-{reduction:.2f}", "%", "gasoil")

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown(
        f'<div class="small-title">Profil de consommation active du gasoil par phase</div>',
        unsafe_allow_html=True
    )

    col_g1, col_g2 = st.columns(2)

    with col_g1:
        fig_anc_energy = go.Figure()

        y_anc = anc_gas["Conso_gasoil_cumulee_phase_m3"].copy()
        phase_anc = anc_gas["Phase"].astype(str).str.lower()

        mask_sans_gasoil_anc = (
            phase_anc.str.contains("manoeuvre", na=False)
            | phase_anc.str.contains("manœuvre", na=False)
            | phase_anc.str.contains("soufflage", na=False)
            | phase_anc.str.contains("derniere", na=False)
            | phase_anc.str.contains("fin", na=False)
        )

        y_anc[mask_sans_gasoil_anc] = 0

        fig_anc_energy.add_trace(go.Scatter(
            x=anc_gas["Temps_global_h"],
            y=y_anc,
            mode="lines",
            name="Ancien procédé",
            line=dict(color=RED, width=3),
            fill="tozeroy"
        ))

        fig_anc_energy.update_layout(
            title="Ancien procédé",
            height=360,
            plot_bgcolor="white",
            paper_bgcolor="white",
            xaxis_title="Temps global (h)",
            yaxis_title="Consommation active (m³)",
            margin=dict(l=45, r=20, t=45, b=45),
            font=dict(family="Inter", color=TEXT)
        )

        st.plotly_chart(fig_anc_energy, use_container_width=True)

    with col_g2:
        fig_new_energy = go.Figure()

        y_new = new_gas["Conso_gasoil_cumulee_phase_m3"].copy()
        phase_new = new_gas["Phase"].astype(str).str.lower()

        mask_sans_gasoil_new = (
            phase_new.str.contains("manoeuvre", na=False)
            | phase_new.str.contains("manœuvre", na=False)
            | phase_new.str.contains("soufflage", na=False)
            | phase_new.str.contains("derniere", na=False)
            | phase_new.str.contains("fin", na=False)
        )

        y_new[mask_sans_gasoil_new] = 0

        fig_new_energy.add_trace(go.Scatter(
            x=new_gas["Temps_h"],
            y=y_new,
            mode="lines",
            name="Nouvelle solution",
            line=dict(color=BLUE, width=3),
            fill="tozeroy"
        ))

        fig_new_energy.update_layout(
            title="Nouvelle solution",
            height=360,
            plot_bgcolor="white",
            paper_bgcolor="white",
            xaxis_title="Temps (h)",
            yaxis_title="Consommation active (m³)",
            margin=dict(l=45, r=20, t=45, b=45),
            font=dict(family="Inter", color=TEXT)
        )

        st.plotly_chart(fig_new_energy, use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)

    col_table, col_bar = st.columns([1.2, 1])

    with col_table:
        st.markdown(
            f'<div class="small-title">Consommation par phase</div>',
            unsafe_allow_html=True
        )

        df_phase = pd.DataFrame({
            "Phase": [
                "Chauffage 1",
                "Manœuvres / Soufflage",
                "Chauffage 2",
                "Manœuvres finales",
                "TOTAL"
            ],
            "Ancien (m³)": [
                ancien_ch1,
                0.00,
                ancien_ch2,
                0.00,
                ancien_total
            ],
            "Nouveau (m³)": [
                nouveau_ch1,
                0.00,
                nouveau_ch2,
                0.00,
                nouveau_total
            ],
            "Économie (m³)": [
                ancien_ch1 - nouveau_ch1,
                0.00,
                ancien_ch2 - nouveau_ch2,
                0.00,
                economie
            ]
        })

        st.dataframe(df_phase, use_container_width=True, hide_index=True)

    with col_bar:
        st.markdown(
            f'<div class="small-title">Comparaison globale</div>',
            unsafe_allow_html=True
        )

        fig_bar = go.Figure()

        fig_bar.add_trace(go.Bar(
            x=["Ancien procédé", "Nouvelle solution"],
            y=[ancien_total, nouveau_total],
            marker_color=[RED, BLUE],
            text=[
                f"{ancien_total:.2f} m³",
                f"{nouveau_total:.2f} m³"
            ],
            textposition="outside",
            textfont=dict(
                size=17,
                color="#111827"
            ),
            cliponaxis=False
        ))

        fig_bar.add_annotation(
            x=0.69,
            y=53,
            text=f"-{reduction:.2f} %",
            showarrow=True,
            arrowhead=3,
            arrowsize=1.2,
            arrowwidth=2,
            arrowcolor=OCP,
            ax=0,
            ay=-55,
            font=dict(
                size=25,
                color=OCP,
                family="Inter"
            )
        )

        fig_bar.update_layout(
            height=440,
            showlegend=False,
            plot_bgcolor="white",
            paper_bgcolor="white",
            margin=dict(
                l=60,
                r=30,
                t=90,
                b=60
            ),
            font=dict(
                family="Inter",
                color=TEXT
            ),
            yaxis=dict(
                title="Consommation (m³)",
                range=[0, 90],
                gridcolor="rgba(148,163,184,0.25)",
                zeroline=False
            ),
            xaxis=dict(
                showgrid=False
            )
        )

        st.plotly_chart(fig_bar, use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown(f"""
    <div style="
        background:#E8F3FF;
        border-radius:14px;
        padding:18px 22px;
        color:#0057A8;
        font-size:16px;
        line-height:1.8;
        margin-top:8px;
    ">
        <b>Points clés :</b><br>
        ✅ La nouvelle solution réduit la consommation de gasoil de <b>{reduction:.2f} %</b> par démarrage.<br>
        ✅ L’économie directe de combustible est de <b>{economie:.2f} m³</b> par démarrage.<br>
        ✅ L'utilisation du combustible est concentrée exclusivement sur les séquences de chauffage, garantissant une absence totale de consommation durant les phases de manœuvres et de soufflage.<br>
        ✅ Cette réduction constitue la base du gain économique présenté dans la section Analyse économique.
    </div>
    """, unsafe_allow_html=True)

elif page == 'Analyse Économique':

    prix_gasoil_litre = 14.52761

    gasoil_ancien_total = 70.00
    gasoil_nouveau_total = 30.11
    gasoil_gain = gasoil_ancien_total - gasoil_nouveau_total

    ancien_ch1 = 38.80
    ancien_ch2 = 31.20

    nouveau_ch1 = 9.96
    nouveau_ch2 = 20.15

    gain_h2so4 = 1848.44
    gain_p2o5 = 606.05
    gain_vhp = 2144.19
    gain_mwh = 357.36

    cout_ancien = gasoil_ancien_total * 1000 * prix_gasoil_litre
    cout_nouveau = gasoil_nouveau_total * 1000 * prix_gasoil_litre
    gain_dh = cout_ancien - cout_nouveau
    reduction_pct = gasoil_gain / gasoil_ancien_total * 100

    def dh(x):
        return f"{x:,.2f}".replace(",", " ").replace(".", ",")

    st.markdown(
        f"""
        <h2 style="color:{NAVY}; margin-bottom:4px;">
            Analyse économique — Nouveau procédé vs Ancien procédé
        </h2>
        <p style="color:{MUTED}; font-size:15px;">
            Évaluation du gain économique direct lié à la réduction de la consommation de gasoil
            et des impacts industriels associés.
        </p>
        """,
        unsafe_allow_html=True
    )

    c1, c2, c3, c4, c5 = st.columns(5)

    with c1:
        kpi("Gain économique", dh(gain_dh), "DH / démarrage", "gasoil")
    with c2:
        kpi("Gasoil économisé", f"{gasoil_gain:.2f}", "m³ / démarrage", "gasoil")
    with c3:
        kpi("Gain temps", "25 h 21 min", "par démarrage", "time")
    with c4:
        kpi("H₂SO₄ récupéré", f"{gain_h2so4:.2f}", "TMH / démarrage", "h2so4")
    with c5:
        kpi("Électricité", f"{gain_mwh:.2f}", "MWh équivalent", "elec")

    st.markdown("<br>", unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1.25, 1, 1])

    with col1:
        fig_cost = go.Figure()

        fig_cost.add_trace(
            go.Bar(
                x=["Ancien procédé", "Nouveau procédé"],
                y=[cout_ancien, cout_nouveau],
                marker_color=[NAVY, OCP],
                text=[
                    f"{dh(cout_ancien)} DH",
                    f"{dh(cout_nouveau)} DH"
                ],
                textposition="outside",
                textfont=dict(size=16, color=NAVY, family="Inter"),
                width=0.55,
                cliponaxis=False
            )
        )

        fig_cost.add_shape(
            type="line",
            x0=0.35,
            x1=1,
            y0=cout_ancien * 1.08,
            y1=cout_ancien * 1.08,
            xref="x",
            yref="y",
            line=dict(color="#64748B", width=2, dash="dash")
        )

        fig_cost.add_annotation(
            x=1,
            y=cout_nouveau * 1.10,
            ax=1,
            ay=cout_ancien * 0.98,
            xref="x",
            yref="y",
            axref="x",
            ayref="y",
            showarrow=True,
            arrowhead=3,
            arrowsize=1.2,
            arrowwidth=2.5,
            arrowcolor="#16A34A"
        )

        fig_cost.add_annotation(
            x=1,
            y=(cout_ancien + cout_nouveau) / 2,
            text=f"<b>-{reduction_pct:.2f} %</b>",
            showarrow=False,
            font=dict(size=30, color="#16A34A", family="Inter"),
            bgcolor="rgba(255,255,255,0.85)"
        )

        fig_cost.update_layout(
            title=dict(
                text="1. Coût total gasoil par démarrage",
                font=dict(size=18, color=NAVY, family="Inter")
            ),
            height=500,
            paper_bgcolor="white",
            plot_bgcolor="white",
            showlegend=False,
            margin=dict(l=55, r=25, t=75, b=60),
            font=dict(family="Inter", color=TEXT),
            xaxis=dict(
                title="",
                showgrid=False,
                zeroline=False,
                tickfont=dict(size=13, color=MUTED)
            ),
            yaxis=dict(
                title="Coût (DH)",
                range=[0, cout_ancien * 1.35],
                gridcolor="rgba(0,0,0,0.08)",
                zeroline=False,
                tickformat=".2s",
                tickfont=dict(size=12, color=MUTED)
            )
        )

        st.plotly_chart(fig_cost, use_container_width=True)

    with col2:
        fig_a = go.Figure(go.Pie(
            labels=["Chauffage 1", "Chauffage 2"],
            values=[ancien_ch1, ancien_ch2],
            hole=0.58,
            marker=dict(colors=[BLUE, OCP]),
            texttemplate="%{value:.1f} m³<br>%{percent}",
            textposition="inside"
        ))

        fig_a.update_layout(
            title="2. Ancien procédé<br><span style='font-size:13px'>Total : 70.00 m³</span>",
            height=390,
            margin=dict(l=10, r=10, t=60, b=20),
            paper_bgcolor="white",
            showlegend=True,
            legend=dict(orientation="h", y=-0.05, x=0.5, xanchor="center")
        )

        st.plotly_chart(fig_a, use_container_width=True)

    with col3:
        fig_n = go.Figure(go.Pie(
            labels=["Chauffage vapeur-air", "Chauffage 2"],
            values=[nouveau_ch1, nouveau_ch2],
            hole=0.58,
            marker=dict(colors=[BLUE, OCP]),
            texttemplate="%{value:.2f} m³<br>%{percent}",
            textposition="inside"
        ))

        fig_n.update_layout(
            title="3. Nouveau procédé<br><span style='font-size:13px'>Total : 30.11 m³</span>",
            height=390,
            margin=dict(l=10, r=10, t=60, b=20),
            paper_bgcolor="white",
            showlegend=True,
            legend=dict(orientation="h", y=-0.05, x=0.5, xanchor="center")
        )

        st.plotly_chart(fig_n, use_container_width=True)

    col4, col5 = st.columns([1, 1.45])

    with col4:
    
      st.markdown("### 4. Impact industriel récupéré")

      impact_data = [
        ("🧪", "Production H₂SO₄ récupérée", f"{gain_h2so4:.2f} TMH", "#7C3AED"),
        ("📦", "Production P₂O₅ équivalente", f"{gain_p2o5:.2f} T", BLUE),
        ("☁️", "Vapeur HP disponible", f"{gain_vhp:.2f} T", BLUE),
        ("⚡", "Électricité équivalente", f"{gain_mwh:.2f} MWh", OCP),
      ]

      for icon, label, value, color in impact_data:
        c_icon, c_text, c_value = st.columns([0.25, 1.35, 0.95])

        with c_icon:
            st.markdown(f"<div style='font-size:22px;'>{icon}</div>", unsafe_allow_html=True)

        with c_text:
            st.markdown(
                f"<div style='font-size:15px;font-weight:700;margin-top:4px;'>{label}</div>",
                unsafe_allow_html=True
            )

        with c_value:
            st.markdown(
                f"<div style='font-size:20px;font-weight:800;color:{color};text-align:right;'>{value}</div>",
                unsafe_allow_html=True
            )

        st.markdown(
            "<hr style='margin:8px 0;border:0;border-top:1px solid #E5E7EB;'>",
            unsafe_allow_html=True
        )
    with col5:
        prix_range = np.array([10, 12, 14, 14.52761, 16, 18, 20])
        gain_sens = gasoil_gain * 1000 * prix_range

        fig_sens = go.Figure()
        text_sens = [f"{v:,.0f}".replace(",", " ") for v in gain_sens]

        text_pos = [
         "bottom right",    # 10
         "bottom center",   # 12
         "bottom center",   # 14
         "top center",      # 14.52761 (prix réel)
         "bottom center",   # 16
         "bottom center",   # 18
         "bottom left"      # 20
        ]

        fig_sens.add_trace(go.Scatter(
          x=prix_range,
          y=gain_sens,
          mode="lines+markers+text",
          text=text_sens,
          textposition=text_pos,
          textfont=dict(size=11, color=NAVY),
          line=dict(color=BLUE, width=3),
          marker=dict(size=8, color=OCP),
          fill="tozeroy",
          fillcolor="rgba(37,99,235,0.10)",
          name="Gain économique"
        ))
        fig_sens.add_vline(
          x=prix_gasoil_litre,
          line_dash="dash",
          line_color=ORANGE
        )

        

        fig_sens.update_layout(
            title="5. Sensibilité du gain économique au prix du gasoil",
            xaxis_title="Prix gasoil (DH/L)",
            yaxis_title="Gain économique direct (DH / démarrage)",
            height=420,
            margin=dict(l=60, r=60, t=60, b=45),
            paper_bgcolor="white",
            plot_bgcolor="white",
            font=dict(family="Inter", color=TEXT),
            xaxis=dict(gridcolor=GRID, zeroline=False,range=[9.4, 20.6]),
            yaxis=dict(gridcolor=GRID, zeroline=False)
        )

        st.plotly_chart(fig_sens, use_container_width=True)

    st.info(
        "Les gains économiques directs sont calculés uniquement sur l’économie de gasoil. "
        "Les gains de production et d’énergie sont présentés comme impacts industriels récupérés, "
        "sans valorisation financière additionnelle faute de prix unitaires validés."
    )
st.markdown('<div class="footer">Source : Données industrielles OCP Safi — Évaluation comparative des stratégies de démarrage</div>', unsafe_allow_html=True)