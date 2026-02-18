import streamlit as st
import pandas as pd
import os
import io

# --- 1. CONFIGURATION ---
st.set_page_config(page_title="OCP - Dashboard S&OE Expert", layout="wide")

# --- 2. FONCTION DE CHARGEMENT UNIVERSELLE ---
@st.cache_data
def charger_data(file, sheet):
    try:
        df_raw = pd.read_excel(file, sheet_name=sheet, header=None, engine='openpyxl')
        
        # 1. On cherche la ligne qui contient VRAIMENT les dates (ex: 01/01/26)
        # On scanne les 20 premières lignes pour être sûr
        idx_dates = None
        for i in range(min(20, len(df_raw))):
            row_values = df_raw.iloc[i].astype(str)
            if row_values.str.contains(r'\d{2}/\d{2}', regex=True).any():
                idx_dates = i
                break
        
        if idx_dates is None:
            return pd.DataFrame()

        # 2. On définit les colonnes en utilisant la ligne de dates trouvée
        header_row = df_raw.iloc[idx_dates].fillna("").tolist()
        
        new_cols = []
        for j, val in enumerate(header_row):
            val_str = str(val).strip()
            # Si la cellule est vide ou contient 'nan', on lui donne un nom Info_X
            if val_str == "" or "nan" in val_str.lower():
                new_cols.append(f"Info_{j}")
            else:
                # Si c'est une date, on nettoie pour n'avoir que le JJ/MM/AA
                new_cols.append(val_str.split(" ")[0])

        # 3. Création du DataFrame à partir de la ligne juste après les dates
        df = df_raw.iloc[idx_dates + 1:].copy()
        df.columns = new_cols
        
        # 4. Nettoyage spécifique Production Planning (Colonnes A et B et Titre)
        if sheet == "ProductionPlanning":
            # Supprimer les colonnes A et B (indices 0 et 1)
            df = df.iloc[:, 2:]
            # Forcer le titre demandé à la ligne 0, colonne Info_2 (qui est maintenant à l'index 0 ou 1)
            if not df.empty:
                # On cherche la première colonne qui n'est pas une date
                cols_info = [c for c in df.columns if "Info" in c]
                if len(cols_info) > 1:
                    df.iloc[0, df.columns.get_loc(cols_info[1])] = "atterissage ACP 29"

        # 5. Remplissage des cellules fusionnées (ffill)
        df.iloc[:, 0:10] = df.iloc[:, 0:10].ffill()
        
        return df.dropna(how='all', axis=0).reset_index(drop=True)
    except Exception as e:
        st.error(f"Erreur sur la feuille {sheet} : {e}")
        return pd.DataFrame()

# --- 3. LOGIQUE PRINCIPALE ---
uploaded_file = st.sidebar.file_uploader("Mettre à jour Excel", type=["xlsm", "xlsx"])
local_file = "Inventory Projection.xlsm"
source = uploaded_file if uploaded_file else (local_file if os.path.exists(local_file) else None)

choix_feuille = st.sidebar.radio("Sélectionner la feuille :", ["ProductionPlanning", "ACP", "ACS"])

if source:
    df = charger_data(source, choix_feuille)
    
    if not df.empty:
        # Identification dynamique des colonnes (Dates vs Infos)
        cols_dates = [c for c in df.columns if any(char.isdigit() for char in c) and ("/" in c or "-" in c)]
        cols_info = [c for c in df.columns if c not in cols_dates]

        # --- FILTRES ---
        st.sidebar.subheader("Filtres")
        df_filtered = df.copy()
        
        # Filtre par Info (Dynamique selon la feuille)
        for col in cols_info[:3]:
            opts = sorted(df[col].unique().astype(str))
            sel = st.sidebar.multiselect(f"Filtrer {col}", opts)
            if sel:
                df_filtered = df_filtered[df_filtered[col].astype(str).isin(sel)]
        
        # Filtre par Date
        sel_dates = st.sidebar.multiselect("📅 Choisir Dates", cols_dates)
        show_dates = sel_dates if sel_dates else cols_dates

        # --- AFFICHAGE ---
        st.subheader(f"Vue : {choix_feuille}")
        st.dataframe(df_filtered[cols_info + show_dates], use_container_width=True, height=600)
else:
    st.info("Veuillez charger votre fichier Excel.")
