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
        
        if sheet == "ProductionPlanning":
            idx_dates = 0
            for i in range(len(df)):
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
            df_data = df_data.iloc[:, 2:]
            
            if not df_data.empty:
                col_target = "Info_3" if "Info_3" in df_data.columns else df_data.columns[1]
                df_data[col_target] = df_data[col_target].astype(object)
                df_data.loc[1:, col_target] = None 
                df_data.loc[0, col_target] = "atterissage ACP 29"
            
            return df_data.reset_index(drop=True)

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

# --- AJOUT DU BOUTON SPÉCIFIQUE ---
st.sidebar.markdown("---")
btn_atterissage = st.sidebar.toggle("🎯 Focus Atterissage ACP 29")

# --- 5. LOGIQUE PRINCIPALE ---
if source:
    df_brut = charger_data(source, choix_feuille)
    if not df_brut.empty:
        df = df_brut.copy()
        
        dates_disponibles = [c for c in df.columns if any(char.isdigit() for char in str(c)) and ('/' in str(c) or '-' in str(c))]
        cols_infos = [c for c in df.columns if c not in dates_disponibles]
        
        for d_col in dates_disponibles:
            df[d_col] = pd.to_numeric(df[d_col], errors='coerce').fillna(0)

        # --- LOGIQUE DU BOUTON ATTERISSAGE ---
        if btn_atterissage and choix_feuille == "ProductionPlanning":
            # On ne garde que la première ligne
            df_focus = df.iloc[[0]].copy()
            # On filtre pour ne garder que les colonnes info + les dates avec une valeur > 0
            cols_actives = [c for c in dates_disponibles if df_focus[c].sum() > 0]
            df_affichage = df_focus[cols_infos + cols_actives]
            st.success("Affichage uniquement de la ligne d'atterrissage (colonnes actives)")
        else:
            selection_dates = st.sidebar.multiselect("📅 Filtrer par Date(s) :", dates_disponibles)
            for col in cols_infos[:3]:
                if col in df.columns:
                    options = sorted(df[col].astype(str).unique())
                    selection = st.sidebar.multiselect(f"Sélectionner {col} :", options)
                    if selection: df = df[df[col].astype(str).isin(selection)]
            
            dates_a_afficher = selection_dates if selection_dates else dates_disponibles
            df_affichage = df[cols_infos + dates_a_afficher]

        # --- AFFICHAGE ---
        st.dataframe(df_affichage, use_container_width=True, height=400)

        # Export
        output = io.BytesIO()
        df_affichage.to_excel(output, index=False)
        st.sidebar.download_button("📥 Télécharger Rapport", data=output.getvalue(), file_name=f"Rapport_{choix_feuille}.xlsx")
else:
    st.info("👋 Bonjour Monsieur Adil, veuillez charger un fichier Excel.")
