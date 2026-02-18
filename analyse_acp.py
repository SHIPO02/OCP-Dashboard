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

# --- 3. FONCTION DE CHARGEMENT INTELLIGENTE ---
@st.cache_data
def charger_data(file, sheet):
    try:
        df_raw = pd.read_excel(file, sheet_name=sheet, header=None, engine='openpyxl')
        
        # 1. On cherche la ligne des dates (format date Excel ou texte avec '/')
        idx_dates = 0
        for i in range(min(10, len(df_raw))):
            if df_raw.iloc[i].astype(str).str.contains(r'\d{2}/\d{2}', regex=True).any():
                idx_dates = i
                break
        
        # 2. On récupère les lignes d'en-tête (Dates + Sous-titres)
        row_dates = df_raw.iloc[idx_dates].fillna("")
        row_sub = df_raw.iloc[idx_dates + 1].fillna("")
        
        # 3. Création des noms de colonnes propres
        final_cols = []
        for d, s in zip(row_dates, row_sub):
            d_str, s_str = str(d).strip(), str(s).strip()
            if d_str == "" or "nan" in d_str.lower():
                name = s_str if s_str != "" else "Info"
            else:
                # Si c'est une date, on garde la date
                name = d_str.split(" ")[0] # Garde juste YYYY-MM-DD ou DD/MM
            final_cols.append(name)

        # Gestion des doublons (Info_1, Info_2...)
        counts = {}
        processed_cols = []
        for c in final_cols:
            if c in counts:
                counts[c] += 1
                processed_cols.append(f"{c}_{counts[c]}")
            else:
                counts[c] = 0
                processed_cols.append(c)

        # 4. Nettoyage des données
        df = df_raw.iloc[idx_dates + 2:].copy()
        df.columns = processed_cols
        
        # Remplissage des cellules fusionnées (OIJ Old, etc.)
        # On remplit les 6 premières colonnes qui sont généralement les colonnes de description
        df.iloc[:, 0:6] = df.iloc[:, 0:6].ffill()
        
        # Supprimer les lignes totalement vides
        df = df.dropna(how='all', axis=0)
        
        return df.reset_index(drop=True)
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
        # Séparer les colonnes Info (description) et les colonnes Dates (données numériques)
        cols_info = [c for c in df.columns if any(x in c for x in ["Info", "Product", "Volume", "Unit", "Line"])]
        cols_dates = [c for c in df.columns if c not in cols_info]

        # --- FILTRES ---
        st.sidebar.markdown("---")
        st.sidebar.subheader("Filtres")
        
        df_filtered = df.copy()
        # On crée des filtres pour les 3 premières colonnes d'info
        for col in cols_info[:3]:
            options = sorted(df[col].unique().astype(str))
            selection = st.sidebar.multiselect(f"Filtrer par {col}", options)
            if selection:
                df_filtered = df_filtered[df_filtered[col].astype(str).isin(selection)]

        # --- AFFICHAGE ---
        st.subheader(f"Tableau : {choix_feuille}")
        
        # Metrics simples
        c1, c2 = st.columns(2)
        with c1:
            total = pd.to_numeric(df_filtered[cols_dates].stack(), errors='coerce').sum()
            st.metric("Volume Total (Sélection)", f"{total:,.0f} T")
        with c2:
            st.metric("Lignes affichées", len(df_filtered))

        # Le tableau complet
        st.dataframe(df_filtered, use_container_width=True, height=600)

        # Export
        output = io.BytesIO()
        df_filtered.to_excel(output, index=False)
        st.sidebar.download_button("📥 Télécharger ce tableau", data=output.getvalue(), file_name="export_ocp.xlsx")
else:
    st.info("Veuillez charger le fichier 'Inventory Projection.xlsm' pour voir les données.")
