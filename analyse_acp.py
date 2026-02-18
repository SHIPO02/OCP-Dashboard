import streamlit as st
import pandas as pd
import os
import io

# --- 1. CONFIGURATION ---
st.set_page_config(page_title="OCP - Dashboard S&OE Expert", layout="wide")

st.markdown("""
    <style>
    [data-testid="stSidebar"] { background-color: #83B81A; }
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
    if os.path.exists("logo_ocp.png"): st.image("logo_ocp.png", width=140)
with col_titre:
    st.title("OCP Project")
    st.subheader("Bonjour Monsieur Adil Elbarmaqui")

# --- 3. FONCTION DE CHARGEMENT ---
@st.cache_data
def charger_data(file, sheet):
    try:
        df_raw = pd.read_excel(file, sheet_name=sheet, header=None, engine='openpyxl')
        
        # 1. Identifier la ligne des dates
        idx_dates = 0
        for i in range(min(15, len(df_raw))):
            if df_raw.iloc[i].astype(str).str.contains(r'\d{2}/\d{2}', regex=True).any():
                idx_dates = i
                break
        
        # 2. Préparation des colonnes
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

        counts = {}
        processed_cols = []
        for c in final_cols:
            if c in counts:
                counts[c] += 1
                processed_cols.append(f"{c}_{counts[c]}")
            else:
                counts[c] = 0
                processed_cols.append(c)

        # 3. Création du DataFrame et ffill
        df = df_raw.iloc[idx_dates + 2:].copy()
        df.columns = processed_cols
        df.iloc[:, 0:10] = df.iloc[:, 0:10].ffill()
        
        # 4. SUPPRESSION DES COLONNES A ET B (Indices 0 et 1)
        df = df.iloc[:, 2:]
        
        # 5. FILTRE STRICT SUR "atterissage ACP 29"
        # On cherche dans toutes les colonnes d'info si le texte existe
        mask = df.astype(str).apply(lambda x: x.str.contains("atterissage ACP 29", case=False, na=False)).any(axis=1)
        df = df[mask]

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
    df_affichage = charger_data(source, choix_feuille)
    
    if not df_affichage.empty:
        # Nettoyage visuel : retirer les colonnes qui ne sont que des "None" ou vides
        df_affichage = df_affichage.loc[:, (df_affichage != "None").any(axis=0)]
        
        st.subheader(f"Planning : {choix_feuille}")
        st.write("Affichage filtré sur : **atterissage ACP 29**")
        
        # Affichage
        st.dataframe(df_affichage, use_container_width=True, height=600)

        # Export
        output = io.BytesIO()
        df_affichage.to_excel(output, index=False)
        st.sidebar.download_button("📥 Télécharger Rapport", data=output.getvalue(), file_name="Rapport_ACP29.xlsx")
else:
    st.info("Veuillez charger le fichier Excel.")
