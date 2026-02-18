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
.header-vert {
    background-color: #83B81A;
    color: white;
    padding: 12px;
    border-radius: 8px 8px 0 0;
    font-weight: bold;
    text-align: center;
    font-size: 18px;
    margin-top: 20px;
}
.spacer { margin-bottom: 30px; }
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

# --- 3. FONCTION DE NETTOYAGE STRICT ---
def filtrer_colonnes_vides(df_row):
    """Supprime toutes les colonnes qui sont vides, None, ou égales à 0 sur une ligne donnée."""
    def est_valide(val):
        if pd.isna(val) or val is None: return False
        s_val = str(val).strip().lower()
        if s_val in ["", "none", "nan", "info", "0", "0.0"]: return False
        try:
            if float(val) == 0: return False
        except:
            pass
        return True

    cols_utiles = [c for c in df_row.columns if est_valide(df_row[c].iloc[0])]
    return df_row[cols_utiles]

# --- 4. CHARGEMENT DES DONNÉES ---
@st.cache_data
def charger_data(file, sheet):
    try:
        df = pd.read_excel(file, sheet_name=sheet, header=None, engine='openpyxl')
        df = df.dropna(how='all', axis=0).dropna(how='all', axis=1).reset_index(drop=True)
        
        # Détection de la ligne d'en-tête (Dates JJ/MM/AA)
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

        # Spécifique Production Planning (Suppression A et B)
        if sheet == "ProductionPlanning":
            df_data = df_data.iloc[:, 2:]
            if not df_data.empty:
                cols_inf = [c for c in df_data.columns if "Info" in c]
                target = cols_inf[1] if len(cols_inf) > 1 else df_data.columns[0]
                df_data.loc[0, target] = "atterissage ACP 29"
        
        return df_data.reset_index(drop=True)
    except Exception as e:
        return pd.DataFrame()

# --- 5. INTERFACE SIDEBAR ---
st.sidebar.title("🚀 Pilotage S&OE")
uploaded_file = st.sidebar.file_uploader("Mettre à jour Excel", type=["xlsm", "xlsx"])

dossier = os.path.dirname(os.path.abspath(__file__))
local_file = os.path.join(dossier, "Inventory Projection.xlsm")
source = uploaded_file if uploaded_file else (local_file if os.path.exists(local_file) else None)

st.sidebar.markdown("---")
btn_global = st.sidebar.toggle("🎯 Focus Atterrissage Global")

if not btn_global:
    choix_feuille = st.sidebar.radio("Sélectionner la feuille :", ["ACP", "ACS", "ProductionPlanning"])

# --- 6. LOGIQUE D'AFFICHAGE ---
if source:
    if btn_global:
        # --- MODE GLOBAL : LES 3 FEUILLES NETTOYÉES ---
        for feuille in ["ACP", "ACS", "ProductionPlanning"]:
            df_feuille = charger_data(source, feuille)
            if not df_feuille.empty:
                # On prend la ligne 0 (Atterrissage)
                df_row = df_feuille.iloc[[0]].copy()
                
                # Conversion numérique des colonnes de dates pour le filtrage
                dates_cols = [c for c in df_row.columns if any(char.isdigit() for char in str(c))]
                for c in dates_cols:
                    df_row[c] = pd.to_numeric(df_row[c], errors='coerce').fillna(0)
                
                # Nettoyage strict des colonnes vides/zéro
                df_final = filtrer_colonnes_vides(df_row)
                
                st.markdown(f'<div class="header-vert">ATTERRISSAGE : {feuille}</div>', unsafe_allow_html=True)
                if not df_final.empty:
                    st.table(df_final)
                else:
                    st.warning(f"Aucune donnée active pour la feuille {feuille}")
                st.markdown('<div class="spacer"></div>', unsafe_allow_html=True)
    else:
        # --- MODE NORMAL ---
        df_normal = charger_data(source, choix_feuille)
        if not df_normal.empty:
            st.subheader(f"Vue Détail : {choix_feuille}")
            st.dataframe(df_normal, use_container_width=True)
else:
    st.info("Veuillez charger votre fichier Excel.")

