import streamlit as st
import pandas as pd
import os
import io

# --- 1. CONFIGURATION ET STYLE OCP ---
st.set_page_config(page_title="OCP - Dashboard S&OE Expert", layout="wide")

st.markdown("""
<style>
[data-testid="stSidebar"] { background-color: #83B81A; }
[data-testid="stSidebar"] .stText, [data-testid="stSidebar"] label, [data-testid="stSidebar"] h3, [data-testid="stSidebar"] p {
    color: white !important;
}
.kpi-card {
    background-color: #f8f9fa;
    padding: 20px;
    border-radius: 10px;
    border-left: 5px solid #83B81A;
    text-align: center;
    box-shadow: 2px 2px 5px rgba(0,0,0,0.1);
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

# --- 3. CHARGEMENT DES DONNÉES ---
@st.cache_data
def charger_data(file, sheet):
    try:
        df = pd.read_excel(file, sheet_name=sheet, header=None, engine='openpyxl')
        df = df.dropna(how='all', axis=0).dropna(how='all', axis=1).reset_index(drop=True)
        
        # --- LOGIQUE SPÉCIFIQUE PRODUCTION PLANNING ---
        if sheet == "ProductionPlanning":
            # On cherche la ligne des dates (ex: 01/01/26) qui est généralement en haut
            idx_dates = 0
            for i in range(len(df)):
                if df.iloc[i].astype(str).str.contains(r'\d{2}/\d{2}', regex=True).any():
                    idx_dates = i
                    break
            
            # Récupération des en-têtes (Dates)
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
            
            # On supprime les colonnes A et B (indices 0 et 1)
            df_data = df_data.iloc[:, 2:]
            
            # Injection de "atterissage ACP 29" à la ligne 0
            if not df_data.empty:
                cols_inf = [c for c in df_data.columns if "Info" in c]
                if len(cols_inf) > 1:
                    df_data.iloc[0, df_data.columns.get_loc(cols_inf[1])] = "atterissage ACP 29"
            
            # Remplissage automatique pour les colonnes de gauche
            for c in df_data.columns[:8]:
                df_data[c] = df_data[c].ffill()
            
            return df_data.reset_index(drop=True)

        # --- LOGIQUE ORIGINALE (ACP / ACS) ---
        else:
            headers = df.iloc[1].fillna("Info").tolist()
            df_data = df.iloc[2:].copy()
            new_cols = []
            counts = {}
            for col in headers:
                c_str = str(col).strip()
                if c_str in counts:
                    counts[c_str] += 1
                    new_cols.append(f"{c_str}_{counts[c_str]}")
                else:
                    counts[c_str] = 0
                    new_cols.append(c_str)
            df_data.columns = new_cols
            for c in df_data.columns[:6]: 
                df_data[c] = df_data[c].ffill()
            return df_data

    except Exception as e:
        return pd.DataFrame()

# --- 4. BARRE LATÉRALE ---
st.sidebar.title("🚀 Pilotage S&OE")
uploaded_file = st.sidebar.file_uploader("Mettre à jour Excel", type=["xlsm", "xlsx"])

dossier = os.path.dirname(os.path.abspath(__file__))
local_file = os.path.join(dossier, "Inventory Projection.xlsm")
source = uploaded_file if uploaded_file else (local_file if os.path.exists(local_file) else None)

choix_feuille = st.sidebar.radio("Sélectionner la feuille :", ["ACP", "ACS", "ProductionPlanning"])

# --- 5. LOGIQUE PRINCIPALE ---
if source:
    df_brut = charger_data(source, choix_feuille)
    if not df_brut.empty:
        df = df_brut.copy()
        
        # Identification des dates
        dates_disponibles = [c for c in df.columns if any(char.isdigit() for char in str(c)) and ('/' in str(c) or '-' in str(c))]
        cols_infos = [c for c in df.columns if c not in dates_disponibles]
        
        # Conversion numérique
        for d_col in dates_disponibles:
            df[d_col] = pd.to_numeric(df[d_col], errors='coerce').fillna(0)

        # --- FILTRES SIDEBAR ---
        st.sidebar.markdown("---")
        selection_dates = st.sidebar.multiselect("📅 Filtrer par Date(s) :", dates_disponibles)

        for col in cols_infos[:3]:
            if col in df.columns:
                options = sorted(df[col].astype(str).unique())
                selection = st.sidebar.multiselect(f"Sélectionner {col} :", options)
                if selection: df = df[df[col].astype(str).isin(selection)]

        dates_a_afficher = selection_dates if selection_dates else dates_disponibles
        df_affichage = df[cols_infos + dates_a_afficher]

        # --- CARTES KPI ---
        valeur_totale = df[dates_a_afficher].sum().sum()
        nb_lignes = len(df)
        kpi1, kpi2, kpi3 = st.columns(3)
        with kpi1:
            st.markdown(f"<div class='kpi-card'><h3>Total Mesuré</h3><h2 style='color:#83B81A;'>{valeur_totale:,.2f} T</h2></div>", unsafe_allow_html=True)
        with kpi2:
            st.markdown(f"<div class='kpi-card'><h3>Lignes Actives</h3><h2>{nb_lignes}</h2></div>", unsafe_allow_html=True)
        with kpi3:
            st.markdown(f"<div class='kpi-card'><h3>Périodes</h3><h2>{len(dates_a_afficher)}</h2></div>", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        tab_vue, tab_graph = st.tabs(["📄 Vue Détaillée", "📊 Analyses"])

        with tab_vue:
            st.dataframe(df_affichage, use_container_width=True, height=600)

        with tab_graph:
            if len(dates_a_afficher) > 0:
                st.line_chart(df[dates_a_afficher].T.sum(axis=1))

        # Export
        output = io.BytesIO()
        df_affichage.to_excel(output, index=False)
        st.sidebar.download_button("📥 Télécharger Rapport", data=output.getvalue(), file_name=f"Rapport_{choix_feuille}.xlsx")
else:
    st.info("👋 Bonjour Monsieur Adil, veuillez charger un fichier Excel pour commencer.")
