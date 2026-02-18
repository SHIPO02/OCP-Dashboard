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
        
        # Détection de la ligne d'en-tête (Dates)
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

        # Spécifique Production Planning (Suppression A et B)
        if sheet == "ProductionPlanning":
            df_data = df_data.iloc[:, 2:]
            if not df_data.empty:
                col_target = [c for c in df_data.columns if "Info" in c][1] if len([c for c in df_data.columns if "Info" in c]) > 1 else df_data.columns[0]
                df_data[col_target] = df_data[col_target].astype(object)
                df_data.loc[1:, col_target] = None 
                df_data.loc[0, col_target] = "atterissage ACP 29"
        
        return df_data.reset_index(drop=True)
    except Exception as e:
        return pd.DataFrame()

# --- 4. BARRE LATÉRALE ---
st.sidebar.title("🚀 Pilotage S&OE")
uploaded_file = st.sidebar.file_uploader("Mettre à jour Excel", type=["xlsm", "xlsx"])

dossier = os.path.dirname(os.path.abspath(__file__))
local_file = os.path.join(dossier, "Inventory Projection.xlsm")
source = uploaded_file if uploaded_file else (local_file if os.path.exists(local_file) else None)

choix_feuille = st.sidebar.radio("Sélectionner la feuille :", ["ACP", "ACS", "ProductionPlanning"])

st.sidebar.markdown("---")
btn_focus = st.sidebar.toggle("🎯 Focus Atterrissage (Ligne 1)")

# --- 5. LOGIQUE PRINCIPALE ---
if source:
    df_brut = charger_data(source, choix_feuille)
    if not df_brut.empty:
        df = df_brut.copy()
        
        # Identification des colonnes de dates
        dates_disponibles = [c for c in df.columns if any(char.isdigit() for char in str(c)) and ('/' in str(c) or '-' in str(c))]
        cols_infos = [c for c in df.columns if c not in dates_disponibles]
        
        # Nettoyage numérique des colonnes de dates
        for d_col in dates_disponibles:
            df[d_col] = pd.to_numeric(df[d_col], errors='coerce').fillna(0)

        if btn_focus:
            # --- NETTOYAGE STRICT DES COLONNES VIDES ---
            df_focus = df.iloc[[0]].copy()
            
            # On ne garde que les colonnes où la valeur n'est ni 0, ni None, ni vide
            def est_valide(val):
                s_val = str(val).strip().lower()
                return val != 0 and pd.notna(val) and s_val != "" and s_val != "none" and s_val != "nan"

            cols_a_garder = [c for c in df_focus.columns if est_valide(df_focus[c].iloc[0])]
            
            df_final = df_focus[cols_a_garder]
            
            st.markdown(f'<div class="header-vert">RÉCAPITULATIF ATTERRISSAGE : {choix_feuille}</div>', unsafe_allow_html=True)
            
            if not df_final.empty:
                # Affichage propre sans colonnes inutiles
                st.table(df_final)
            else:
                st.warning("Aucune donnée avec valeur trouvée pour cette ligne.")
        else:
            # Mode normal
            selection_dates = st.sidebar.multiselect("📅 Filtrer par Date(s) :", dates_disponibles)
            for col in cols_infos[:3]:
                if col in df.columns:
                    options = sorted(df[col].astype(str).unique())
                    selection = st.sidebar.multiselect(f"Sélectionner {col} :", options)
                    if selection: df = df[df[col].astype(str).isin(selection)]
            
            dates_to_show = selection_dates if selection_dates else dates_disponibles
            df_affichage = df[cols_infos + dates_to_show]
            st.dataframe(df_affichage, use_container_width=True, height=500)

        # Export
        output = io.BytesIO()
        df_export = df_final if btn_focus else df_affichage
        df_export.to_excel(output, index=False)
        st.sidebar.download_button("📥 Télécharger ce tableau", data=output.getvalue(), file_name=f"Focus_{choix_feuille}.xlsx")
else:
    st.info("👋 Veuillez charger un fichier Excel.")

