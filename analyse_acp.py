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
    img_path = "logo_ocp.png"
    if os.path.exists(img_path):
        st.image(img_path, width=140)
with col_titre:
    st.title("OCP Project")
    st.subheader("Bonjour Monsieur Adil Elbarmaqui")

st.markdown("---")


# --- 3. CHARGEMENT DES DONNÉES (VERSION CORRIGÉE) ---
@st.cache_data
def charger_data(file, sheet):
    try:
        # Lecture brute
        df = pd.read_excel(file, sheet_name=sheet, header=None, engine='openpyxl')
        df = df.dropna(how='all', axis=0).dropna(how='all', axis=1).reset_index(drop=True)

        # Détection de la ligne d'en-tête (cherche la ligne avec des dates ou 'Production plan')
        header_row = 0
        for i in range(len(df)):
            row_str = df.iloc[i].astype(str).values
            if any("01/" in s or "Production plan" in s for s in row_str):
                header_row = i
                break
        
        # Définition des colonnes
        headers = df.iloc[header_row].fillna("Info").tolist()
        df_data = df.iloc[header_row + 1:].copy()
        
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
        
        # REMPLISSAGE DES CELLULES FUSIONNÉES (Crucial pour l'affichage complet)
        cols_infos = [c for c in df_data.columns if "Info" in c or "Production" in c or "OIJ" in str(c)]
        df_data[cols_infos] = df_data[cols_infos].ffill()
        
        return df_data.reset_index(drop=True)
    except Exception as e:
        st.error(f"Erreur de lecture : {e}")
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
    df = charger_data(source, choix_feuille)
    if not df.empty:
        # Séparation Infos vs Dates
        cols_infos = [c for c in df.columns if any(x in str(c) for x in ["Info", "Production", "Total", "Product", "Volume"])]
        dates_disponibles = [c for c in df.columns if c not in cols_infos]
        
        # Nettoyage numérique des dates
        for d in dates_disponibles:
            df[d] = pd.to_numeric(df[d], errors='coerce').fillna(0)

        # Filtres dynamiques
        st.sidebar.markdown("---")
        for col in cols_infos[:3]:
            options = sorted(df[col].astype(str).unique())
            sel = st.sidebar.multiselect(f"Filtrer {col}", options)
            if sel: df = df[df[col].astype(str).isin(sel)]

        selection_dates = st.sidebar.multiselect("📅 Sélection Dates :", dates_disponibles)
        dates_a_afficher = selection_dates if selection_dates else dates_disponibles
        
        df_affichage = df[cols_infos + dates_a_afficher]

        # KPIs
        val_tot = df[dates_a_afficher].sum().sum()
        k1, k2 = st.columns(2)
        k1.metric("Total Mesuré", f"{val_tot:,.0f} T")
        k2.metric("Lignes affichées", len(df))

        tab_vue, tab_graph = st.tabs(["📄 Vue Détaillée", "📊 Analyses"])
        with tab_vue:
            st.dataframe(df_affichage, use_container_width=True, height=600)
        
        with tab_graph:
            st.info("Sélectionnez des données pour générer les graphiques.")
            if len(dates_a_afficher) > 0:
                st.line_chart(df[dates_a_afficher].T.sum(axis=1))

        # Export
        output = io.BytesIO()
        df_affichage.to_excel(output, index=False)
        st.sidebar.download_button("📥 Télécharger Rapport", data=output.getvalue(), file_name="Rapport_OCP.xlsx")
else:
    st.info("Veuillez charger un fichier Excel.")
