import streamlit as st
import pandas as pd
import os
import io

# --- 1. CONFIGURATION ---
st.set_page_config(page_title="OCP - Dashboard S&OE Expert", layout="wide")

st.markdown("""
    <style>
    [data-testid="stSidebar"] { background-color: #83B81A; }
    .title-ocp { color: #83B81A; font-weight: bold; font-size: 24px; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. EN-TÊTE ---
col_logo, col_titre = st.columns([1, 4])
with col_logo:
    if os.path.exists("logo_ocp.png"): st.image("logo_ocp.png", width=140)
with col_titre:
    st.title("OCP Project")
    st.subheader("Bonjour Monsieur Adil Elbarmaqui")

# --- 3. FONCTION DE CHARGEMENT ---
@st.cache_data
def charger_data(file, sheet):
    try:
        df_raw = pd.read_excel(file, sheet_name=sheet, header=None, engine='openpyxl')
        
        # 1. Identifier la ligne des dates (ex: 01/01/26)
        idx_dates = 0
        for i in range(min(15, len(df_raw))):
            if df_raw.iloc[i].astype(str).str.contains(r'\d{2}/\d{2}', regex=True).any():
                idx_dates = i
                break
        
        # 2. Construction des noms de colonnes
        row_dates = df_raw.iloc[idx_dates].fillna("")
        row_sub = df_raw.iloc[idx_dates + 1].fillna("")
        
        final_cols = []
        for d, s in zip(row_dates, row_sub):
            d_str, s_str = str(d).strip(), str(s).strip()
            if d_str == "" or "nan" in d_str.lower():
                name = s_str if s_str != "" else "Info"
            else:
                name = d_str.split(" ")[0]
            final_cols.append(name)

        # Gérer les doublons de noms (Info_1, Info_2...)
        counts = {}
        processed_cols = []
        for c in final_cols:
            if c in counts:
                counts[c] += 1
                processed_cols.append(f"{c}_{counts[c]}")
            else:
                counts[c] = 0
                processed_cols.append(c)

        # 3. Création du DataFrame
        df = df_raw.iloc[idx_dates + 2:].copy()
        df.columns = processed_cols
        
        # Remplissage des cellules fusionnées (OIJ Old, etc.)
        df.iloc[:, 0:10] = df.iloc[:, 0:10].ffill()
        
        # 4. SUPPRESSION DES COLONNES A ET B (Indices 0 et 1)
        df = df.iloc[:, 2:]
        
        return df.dropna(how='all', axis=0).reset_index(drop=True)
    except Exception as e:
        st.error(f"Erreur : {e}")
        return pd.DataFrame()

# --- 4. LOGIQUE PRINCIPALE ---
uploaded_file = st.sidebar.file_uploader("Mettre à jour Excel", type=["xlsm", "xlsx"])
local_file = "Inventory Projection.xlsm"
source = uploaded_file if uploaded_file else (local_file if os.path.exists(local_file) else None)

choix_feuille = st.sidebar.radio("Sélectionner la feuille :", ["ProductionPlanning", "ACP", "ACS"])

if source:
    df = charger_data(source, choix_feuille)
    
    if not df.empty:
        # Identification des colonnes
        cols_dates = [c for c in df.columns if any(char.isdigit() for char in c) and ("/" in c or "-" in c)]
        cols_info = [c for c in df.columns if c not in cols_dates]

        # --- FILTRES SIDEBAR ---
        st.sidebar.markdown("---")
        st.sidebar.subheader("Filtres")
        
        # Filtre par Date
        sel_dates = st.sidebar.multiselect("📅 Sélectionner Dates", cols_dates)
        dates_to_show = sel_dates if sel_dates else cols_dates
        
        # Filtres par Info (ex: Info_3, Info_4...)
        df_filtered = df.copy()
        for col in cols_info[:3]:
            opts = sorted(df[col].unique().astype(str))
            sel = st.sidebar.multiselect(f"Filtrer {col}", opts)
            if sel:
                df_filtered = df_filtered[df_filtered[col].astype(str).isin(sel)]

        # --- AFFICHAGE ---
        # Affichage du titre demandé comme en-tête de page et non dans le tableau
        st.markdown('<p class="title-ocp">atterissage ACP 29</p>', unsafe_allow_html=True)
        st.info(f"Feuille actuelle : {choix_feuille}")

        # Affichage du tableau final sans les colonnes A et B
        st.dataframe(df_filtered[cols_info + dates_to_show], use_container_width=True, height=600)

        # Export
        output = io.BytesIO()
        df_filtered.to_excel(output, index=False)
        st.sidebar.download_button("📥 Télécharger Rapport", data=output.getvalue(), file_name="Rapport_OCP.xlsx")
else:
    st.info("Veuillez charger le fichier Excel.")

