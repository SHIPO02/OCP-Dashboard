import streamlit as st
import pandas as pd
import os
import io
from datetime import datetime

# --- 1. CONFIGURATION ET STYLE OCP ---
st.set_page_config(page_title="OCP - Dashboard S&OE Expert", layout="wide")

st.markdown("""
<style>
[data-testid="stSidebar"] { background-color: #83B81A; }
[data-testid="stSidebar"] .stText, [data-testid="stSidebar"] label, [data-testid="stSidebar"] h3, [data-testid="stSidebar"] p {
    color: white !important;
}
.header-vert {
    background-color: #83B81A;
    color: white;
    padding: 12px;
    border-radius: 8px 8px 0 0;
    font-weight: bold;
    text-align: center;
    font-size: 18px;
    margin-top: 10px;
}
.kpi-card {
    background-color: #f0f2f6;
    padding: 20px;
    border-radius: 10px;
    border-left: 8px solid #83B81A;
    box-shadow: 2px 2px 5px rgba(0,0,0,0.1);
    text-align: center;
}
.spacer { margin-bottom: 40px; }
</style>
""", unsafe_allow_html=True)

# --- 2. EN-TÊTE ---
col_logo, col_titre = st.columns([1, 4])
with col_logo:
    img_path = "logo_ocp.png" if os.path.exists("logo_ocp.png") else "logo_ocp.png.png"
    if os.path.exists(img_path):
        st.image(img_path, width=140)
with col_titre:
    st.title("OCP Project")
    st.subheader("Bonjour Monsieur Adil Elbarmaqui")

st.markdown("---")

# --- 3. FONCTIONS DE CHARGEMENT ET FILTRAGE ---
@st.cache_data
def charger_data(file, sheet):
    try:
        df = pd.read_excel(file, sheet_name=sheet, header=None, engine='openpyxl')
        df = df.dropna(how='all', axis=0).dropna(how='all', axis=1).reset_index(drop=True)
        
        idx_dates = 1 if sheet != "ProductionPlanning" else 0
        if sheet == "ProductionPlanning":
            for i in range(min(15, len(df))):
                if df.iloc[i].astype(str).str.contains(r'\d{2}/\d{2}', regex=True).any():
                    idx_dates = i
                    break
        
        headers = df.iloc[idx_dates].fillna("Info").tolist()
        df_data = df.iloc[idx_dates + 1:].copy()
        
        new_cols = []
        counts = {}
        for i, col in enumerate(headers):
            c_str = str(col).strip()
            if c_str == "nan" or c_str == "Info": c_str = f"Info_{i}"
            if c_str in counts:
                counts[c_str] += 1
                new_cols.append(f"{c_str}_{counts[c_str]}")
            else:
                counts[c_str] = 0
                new_cols.append(c_str)
        
        df_data.columns = new_cols

        if sheet == "ProductionPlanning":
            df_data = df_data.iloc[:, 2:]
            if not df_data.empty:
                cols_inf = [c for c in df_data.columns if "Info" in c]
                target = cols_inf[1] if len(cols_inf) > 1 else df_data.columns[0]
                df_data[target] = df_data[target].astype(object)
                df_data.loc[1:, target] = None 
                df_data.loc[0, target] = "atterissage ACP 29"
        
        return df_data.reset_index(drop=True)
    except Exception as e:
        return pd.DataFrame()

def est_valide(val):
    s_val = str(val).strip().lower()
    return val != 0 and pd.notna(val) and s_val not in ["", "none", "nan", "info"]

# --- 4. BARRE LATÉRALE ---
st.sidebar.title("🚀 Pilotage S&OE")
uploaded_file = st.sidebar.file_uploader("Mettre à jour Excel", type=["xlsm", "xlsx"])

dossier = os.path.dirname(os.path.abspath(__file__))
local_file = os.path.join(dossier, "Inventory Projection.xlsm")
source = uploaded_file if uploaded_file else (local_file if os.path.exists(local_file) else None)

st.sidebar.markdown("---")
btn_focus = st.sidebar.toggle("🎯 Atterrissage Global")

if not btn_focus:
    choix_feuille = st.sidebar.radio("Sélectionner la feuille :", ["ACP", "ACS", "ProductionPlanning"])
else:
    st.sidebar.info("Vue consolidée activée.")

# --- 5. LOGIQUE PRINCIPALE ---
if source:
    if btn_focus:
        # --- VUE : ATTERRISSAGE GLOBAL ---
        st.subheader("📊 Atterrissage Global")
        for feuille in ["ACP", "ACS", "ProductionPlanning"]:
            df_feuille = charger_data(source, feuille)
            if not df_feuille.empty:
                dates_dispo = [c for c in df_feuille.columns if any(char.isdigit() for char in str(c))]
                for d in dates_dispo:
                    df_feuille[d] = pd.to_numeric(df_feuille[d], errors='coerce').fillna(0)
                df_focus = df_feuille.iloc[[0]].copy()
                cols_utiles = [c for c in df_focus.columns if est_valide(df_focus[c].iloc[0])]
                st.markdown(f'<div class="header-vert">ATTERRISSAGE : {feuille}</div>', unsafe_allow_html=True)
                st.table(df_focus[cols_utiles])
                st.markdown('<div class="spacer"></div>', unsafe_allow_html=True)
    
    else:
        # --- VUE : NORMALE (AVEC FILTRAGE INTELLIGENT POUR ACP) ---
        df_brut = charger_data(source, choix_feuille)
        if not df_brut.empty:
            df = df_brut.copy()
            dates_cols = [c for c in df.columns if any(char.isdigit() for char in str(c)) and ('/' in str(c) or '-' in str(c))]
            info_cols = [c for c in df.columns if c not in dates_cols]
            
            for d_col in dates_cols:
                df[d_col] = pd.to_numeric(df[d_col], errors='coerce').fillna(0)

            # --- LOGIQUE SPÉCIFIQUE ACP : CARTES FILTRAGE ---
            if choix_feuille == "ACP":
                st.sidebar.markdown("---")
                st.sidebar.subheader("Filtres Intelligents ACP")
                
                # 1. Sélection Unité
                unit_col = info_cols[1] if len(info_cols) > 1 else info_cols[0]
                unites = sorted(df[unit_col].dropna().unique().astype(str))
                u_sel = st.sidebar.selectbox("Unité", unites)
                
                # 2. Sélection Mois
                dates_dt = pd.to_datetime(dates_cols, dayfirst=True, errors='coerce')
                mois_dispo = sorted(list(set([d.strftime('%m/%Y') for d in dates_dt if pd.notnull(d)])), 
                                    key=lambda x: datetime.strptime(x, '%m/%Y'))
                m_sel = st.sidebar.selectbox("Mois", mois_dispo)
                
                # 3. Calcul Dynamique
                cols_mois = [dates_cols[i] for i, d in enumerate(dates_dt) if pd.notnull(d) and d.strftime('%m/%Y') == m_sel]
                df_filtre = df[df[unit_col].astype(str) == u_sel]
                total_mois = df_filtre[cols_mois].values.sum()

                # 4. Affichage sous forme de Carte KPI
                c1, c2, c3 = st.columns([1, 2, 1])
                with c2:
                    st.markdown(f"""
                    <div class="kpi-card">
                        <p style="color: #555; font-size: 16px; margin: 0;">Somme Production Mensuelle</p>
                        <h1 style="color: #83B81A; margin: 10px 0;">{total_mois:,.0f} T</h1>
                        <p style="font-weight: bold; margin: 0;">{u_sel} — {m_sel}</p>
                    </div>
                    """, unsafe_allow_html=True)
                st.markdown("<br>", unsafe_allow_html=True)
                
                # Affichage du tableau restreint au mois
                st.dataframe(df_filtre[info_cols + cols_mois], use_container_width=True)
            
            else:
                # --- AFFICHAGE CLASSIQUE ACS / PRODUCTION PLANNING ---
                st.markdown(f"### Détail : {choix_feuille}")
                st.dataframe(df, use_container_width=True, height=500)

            # Export
            output = io.BytesIO()
            df.to_excel(output, index=False)
            st.sidebar.download_button("📥 Télécharger Excel", data=output.getvalue(), file_name=f"{choix_feuille}.xlsx")
else:
    st.info("👋 Veuillez charger un fichier Excel.")
