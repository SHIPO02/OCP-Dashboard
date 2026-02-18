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
.kpi-total {
    background-color: #f0f2f6;
    padding: 15px;
    border-radius: 10px;
    border-left: 5px solid #83B81A;
    margin-bottom: 20px;
}
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

# --- 3. FONCTIONS DE CHARGEMENT ---
@st.cache_data
def charger_data(file, sheet):
    try:
        df = pd.read_excel(file, sheet_name=sheet, header=None, engine='openpyxl')
        df = df.dropna(how='all', axis=0).dropna(how='all', axis=1).reset_index(drop=True)
        
        idx_dates = 0
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
                df_data.loc[0, target] = "atterissage ACP 29"
        
        return df_data.reset_index(drop=True)
    except Exception as e:
        return pd.DataFrame()

# --- 4. BARRE LATÉRALE ---
st.sidebar.title("🚀 Pilotage S&OE")
uploaded_file = st.sidebar.file_uploader("Mettre à jour Excel", type=["xlsm", "xlsx"])

dossier = os.path.dirname(os.path.abspath(__file__))
local_file = os.path.join(dossier, "Inventory Projection.xlsm")
source = uploaded_file if uploaded_file else (local_file if os.path.exists(local_file) else None)

st.sidebar.markdown("---")
btn_global = st.sidebar.toggle("🎯 Atterrissage Global")

# --- 5. LOGIQUE PRINCIPALE ---
if source:
    if btn_global:
        # Code Atterrissage Global (identique au précédent)
        for feuille in ["ACP", "ACS", "ProductionPlanning"]:
            df_f = charger_data(source, feuille)
            if not df_f.empty:
                st.markdown(f'<div class="header-vert">ATTERRISSAGE : {feuille}</div>', unsafe_allow_html=True)
                st.table(df_f.iloc[[0]])
    else:
        choix_feuille = st.sidebar.radio("Sélectionner la feuille :", ["ACP", "ACS", "ProductionPlanning"])
        df = charger_data(source, choix_feuille)
        
        if not df.empty:
            # 1. Identifier les colonnes de dates
            dates_cols = [c for c in df.columns if any(char.isdigit() for char in str(c)) and ('/' in str(c) or '-' in str(c))]
            info_cols = [c for c in df.columns if c not in dates_cols]
            
            # 2. Convertir les dates pour filtrage par mois
            dates_dt = []
            for d in dates_cols:
                try:
                    dates_dt.append(pd.to_datetime(d, dayfirst=True))
                except:
                    dates_dt.append(None)
            
            # 3. Filtres Sidebar pour ACP
            if choix_feuille == "ACP":
                st.sidebar.markdown("---")
                st.sidebar.subheader("Filtrage Production ACP")
                
                # Filtre par Unité
                unit_col = info_cols[1] if len(info_cols) > 1 else info_cols[0]
                unites = sorted(df[unit_col].dropna().unique().astype(str))
                unite_sel = st.sidebar.selectbox("Unités", unites)
                
                # Filtre par Mois
                mois_dispo = sorted(list(set([d.strftime('%m/%Y') for d in dates_dt if d is not None])), key=lambda x: datetime.strptime(x, '%m/%Y'))
                mois_sel = st.sidebar.selectbox("Sélectionner le Mois", mois_dispo)
                
                # Calcul de la somme
                cols_mois = [dates_cols[i] for i, d in enumerate(dates_dt) if d is not None and d.strftime('%m/%Y') == mois_sel]
                df_unit = df[df[unit_col].astype(str) == unite_sel]
                total_prod = pd.to_numeric(df_unit[cols_mois].values.flatten(), errors='coerce').sum()
                
                # Affichage du résultat KPI
                st.markdown(f"""
                <div class="kpi-total">
                    <h3 style='margin:0;'>Total Production {unite_sel}</h3>
                    <h2 style='color:#83B81A; margin:0;'>{total_prod:,.0f} T</h2>
                    <p style='margin:0;'>Période : {mois_sel}</p>
                </div>
                """, unsafe_allow_html=True)
                
                st.dataframe(df_unit[info_cols + cols_mois], use_container_width=True)
            else:
                st.dataframe(df, use_container_width=True)
else:
    st.info("Veuillez charger un fichier Excel.")
