import streamlit as st
import pandas as pd
import numpy as np
import re
import os
import base64
import unicodedata
from io import BytesIO
import xlsxwriter

st.set_page_config(page_title="Audit Sémantique SEO", page_icon="🔍", layout="wide")

def extract_domain(url):
    """Extrait le nom de domaine d'une URL."""
    try:
        if pd.isna(url):
            return np.nan
        match = re.search(r'https?://(?:www\.)?([^/]+)', str(url))
        if match:
            return match.group(1)
        return url
    except:
        return url

def read_ahrefs_file(file):
    """Lit les fichiers d'export Ahrefs avec gestion correcte des types et encodages."""
    try:
        # Pour les fichiers Excel
        if file.name.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(file)
        elif file.name.endswith('.csv'):
            # Détecter si c'est un fichier UTF-16 de Ahrefs
            is_utf16 = False
            try:
                # Lire les premiers octets du fichier pour détecter l'encodage UTF-16
                file_content = file.read(4)
                file.seek(0)  # Remettre le pointeur au début
                
                # Le BOM UTF-16 LE commence par 0xFF 0xFE
                if len(file_content) >= 2 and file_content[0] == 0xFF and file_content[1] == 0xFE:
                    is_utf16 = True
                    print(f"Détecté UTF-16 LE pour {file.name}")
            except:
                pass
                
            try:
                if is_utf16:
                    # Pour les fichiers UTF-16 d'Ahrefs (avec tabulations)
                    df = pd.read_csv(file, encoding='utf-16', sep='\t', engine='python')
                    print(f"Fichier lu en UTF-16 avec tabulations: {file.name}")
                else:
                    # Pour les fichiers UTF-8 standard
                    df = pd.read_csv(file, encoding='utf-8', on_bad_lines='warn')
                    print(f"Fichier lu en UTF-8 avec virgules: {file.name}")
            except Exception as e:
                print(f"Première tentative échouée: {str(e)}")
                # Deuxième tentative avec différents paramètres
                try:
                    df = pd.read_csv(file, encoding='utf-16le', sep='\t', engine='python')
                    print(f"Fichier lu en UTF-16LE avec tabulations: {file.name}")
                except Exception as e2:
                    print(f"Deuxième tentative échouée: {str(e2)}")
                    # Troisième tentative
                    try:
                        df = pd.read_csv(file, encoding='cp1252', engine='python')
                        print(f"Fichier lu en CP1252: {file.name}")
                    except Exception as e3:
                        print(f"Toutes les tentatives ont échoué pour {file.name}")
                        return None
        
        # Afficher les colonnes pour débogage
        print(f"Colonnes dans {file.name}: {df.columns.tolist()}")
        
        # Nettoyage et conversion des types
        if 'Keyword' in df.columns:
            # S'assurer que Keyword est une chaîne
            df['Keyword'] = df['Keyword'].astype(str)
        
        if 'Volume' in df.columns:
            # Convertir Volume en numérique
            df['Volume'] = df['Volume'].astype(str).str.replace(',', '').str.replace(' ', '')
            df['Volume'] = pd.to_numeric(df['Volume'], errors='coerce').fillna(0).astype(int)
        
        if 'Current position' in df.columns:
            # Convertir Position en numérique
            df['Current position'] = pd.to_numeric(df['Current position'], errors='coerce').fillna(1000).astype(int) 
        return df
    except Exception as e:
        print(f"Erreur globale lors de la lecture du fichier {file.name}: {str(e)}")
        return None
            
        for i, encoding in enumerate(encodings_to_try):
                try:
                    sep = separators[i]
                    df = pd.read_csv(file, encoding=encoding, sep=sep, on_bad_lines='warn')
                    print(f"Fichier {file.name} lu avec succès en utilisant l'encodage {encoding} et le séparateur '{sep}'")
                    # Vérifier rapidement si le fichier a été correctement lu en vérifiant le nombre de colonnes
                    if len(df.columns) < 3:
                        print(f"Le fichier {file.name} semble mal formaté avec l'encodage {encoding} (seulement {len(df.columns)} colonnes)")
                        continue
                    break
                except Exception as e:
                    print(f"Échec avec encodage {encoding} pour {file.name}: {str(e)}")
                    continue
            
                    if df is None:
                        print(f"Impossible de lire le fichier {file.name} avec les encodages essayés.")
                    return None
    
        # Afficher les colonnes pour le débogage
        print(f"Colonnes trouvées dans {file.name}: {df.columns.tolist()}")
        
        # Nettoyage et conversion des types
        if 'Keyword' in df.columns:
            # S'assurer que Keyword est une chaîne
            df['Keyword'] = df['Keyword'].astype(str)
        
        if 'Volume' in df.columns:
            # Convertir Volume en numérique
            df['Volume'] = df['Volume'].astype(str).str.replace(',', '').str.replace(' ', '')
            df['Volume'] = pd.to_numeric(df['Volume'], errors='coerce').fillna(0).astype(int)
        
        if 'Current position' in df.columns:
            # Convertir Position en numérique
            df['Current position'] = pd.to_numeric(df['Current position'], errors='coerce').fillna(1000).astype(int)
        
        return df
    except Exception as e:
        print(f"Erreur lors de la lecture du fichier {file.name}: {str(e)}")
        return None
        
        # Nettoyage et conversion des types
        if 'Keyword' in df.columns:
            # S'assurer que Keyword est une chaîne
            df['Keyword'] = df['Keyword'].astype(str)
        
        if 'Volume' in df.columns:
            # Convertir Volume en numérique
            df['Volume'] = df['Volume'].astype(str).str.replace(',', '').str.replace(' ', '')
            df['Volume'] = pd.to_numeric(df['Volume'], errors='coerce').fillna(0).astype(int)
        
        if 'Current position' in df.columns:
            # Convertir Position en numérique
            df['Current position'] = pd.to_numeric(df['Current position'], errors='coerce').fillna(1000).astype(int)
        
        return df
    except Exception as e:
        print(f"Erreur lors de la lecture du fichier: {e}")
        return None

def process_files(uploaded_files, column_mapping):
    """Traite les fichiers pour l'audit sémantique."""
    all_data = []
    source_data = []
    
    with st.spinner("Traitement des fichiers en cours..."):
        for file in uploaded_files:
            # Lire le fichier
            df = read_ahrefs_file(file)
            if df is None or df.empty:
                continue
            
            # Extraire le nom du domaine du nom de fichier
            domain_name = os.path.splitext(file.name)[0]
            
            # Adapter le mapping des colonnes
            rename_dict = {}
            for target_col, source_col in column_mapping.items():
                if source_col in df.columns:
                    rename_dict[source_col] = target_col
                else:
                    # Recherche approximative
                    for col in df.columns:
                        if source_col.lower() in col.lower() or target_col in col.lower():
                            rename_dict[col] = target_col
                            break
            
            # Renommer les colonnes si possible
            if rename_dict:
                df = df.rename(columns=rename_dict)
            
            # Utiliser les premières colonnes si nécessaire
            required_columns = ['keyword', 'volume', 'position', 'current_url']
            missing_required = [col for col in required_columns if col not in df.columns]
            
            if missing_required and len(df.columns) >= 4:
                first_cols = df.columns[:4]
                rename_dict = {first_cols[0]: 'keyword', first_cols[1]: 'volume', 
                                first_cols[2]: 'position', first_cols[3]: 'current_url'}
                df = df.rename(columns=rename_dict)
            
            # S'assurer que les colonnes numériques sont bien numériques
            if 'volume' in df.columns:
                df['volume'] = pd.to_numeric(df['volume'], errors='coerce').fillna(0).astype(int)
            
            if 'position' in df.columns:
                df['position'] = pd.to_numeric(df['position'], errors='coerce').fillna(1000).astype(int)
            
            # Ajouter le domaine
            df['domain'] = domain_name
            
            # Ajouter aux données
            all_data.append(df)
            
            # Copier pour l'onglet source
            source_df = df.copy()
            source_df['source_file'] = file.name
            source_data.append(source_df)
                
    if not all_data:
        return None, None
    
    # Combiner toutes les données
    combined_data = pd.concat(all_data, ignore_index=True)
    source_data_combined = pd.concat(source_data, ignore_index=True)
    
    return combined_data, source_data_combined
def create_competition_audit(combined_data, filters):
    """Crée l'audit de compétition avec les nouveaux filtres améliorés."""
    # Convertir position en numérique (pour être sûr)
    combined_data['position'] = pd.to_numeric(combined_data['position'], errors='coerce')
    combined_data['position'].fillna(1000, inplace=True)
    
    # Regrouper par mot-clé exact (sans normalisation)
    grouped = combined_data.groupby('keyword')
    
    # Obtenir tous les domaines uniques
    domains = sorted(combined_data['domain'].unique())
    
    # Préparer les résultats
    audit_results = []
    
    # Afficher les info de débogage
    st.info(f"Analyse de {len(grouped)} mots-clés sur {len(domains)} domaines")
    
    # Barre de progression
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    # Compter les mots-clés traités
    processed = 0
    total_keywords = len(grouped)
    
    for keyword, group in grouped:
        # Mettre à jour la progression
        processed += 1
        progress = processed / total_keywords
        progress_bar.progress(progress)
        status_text.text(f"Traitement des mots-clés: {processed}/{total_keywords}")
        
        # Nombre de domaines réellement positionnés (position inférieure à 100)
        positioned_domains = group[group['position'] <= 100]['domain'].nunique()
        
        # Appliquer les filtres améliorés selon le type sélectionné
        meets_criteria = False
        
        filter_type = filters['filter_type']
        min_sites = filters['min_sites']
        top_positions = filters['top_positions']
        min_sites_in_top = filters['min_sites_in_top']
        
        # Nombre de domaines positionnés dans le top X
        domains_in_top_x = group[group['position'] <= top_positions]['domain'].nunique()
        
        if filter_type == "Au moins 1 site positionné dans le top 10":
            # Au moins 1 site dans le top 10
            meets_criteria = (group['position'] <= 10).any()
            meets_criteria = meets_criteria and (positioned_domains >= min_sites)
            meets_criteria = meets_criteria and (domains_in_top_x >= min_sites_in_top)
            
        elif filter_type == "Au moins 1 site positionné dans le top 20":
            # Au moins 1 site dans le top 20
            meets_criteria = (group['position'] <= 20).any()
            meets_criteria = meets_criteria and (positioned_domains >= min_sites)
            meets_criteria = meets_criteria and (domains_in_top_x >= min_sites_in_top)
            
        elif filter_type == "Au moins 1 site positionné dans le top 30":
            # Au moins 1 site dans le top 30
            meets_criteria = (group['position'] <= 30).any()
            meets_criteria = meets_criteria and (positioned_domains >= min_sites)
            meets_criteria = meets_criteria and (domains_in_top_x >= min_sites_in_top)
            
        elif filter_type == "Au moins 2 sites positionnés, dont 1 top 10":
            # Au moins 2 sites au total ET au moins 1 dans le top 10
            meets_criteria = (positioned_domains >= 2) and (group['position'] <= 10).any()
            meets_criteria = meets_criteria and (positioned_domains >= min_sites)
            meets_criteria = meets_criteria and (domains_in_top_x >= min_sites_in_top)
            
        elif filter_type == "Au moins 2 sites positionnés, dont 1 top 20":
            # Au moins 2 sites au total ET au moins 1 dans le top 20
            meets_criteria = (positioned_domains >= 2) and (group['position'] <= 20).any()
            meets_criteria = meets_criteria and (positioned_domains >= min_sites)
            meets_criteria = meets_criteria and (domains_in_top_x >= min_sites_in_top)
            
        elif filter_type == "Au moins 2 sites positionnés, dont 1 top 30":
            # Au moins 2 sites au total ET au moins 1 dans le top 30
            meets_criteria = (positioned_domains >= 2) and (group['position'] <= 30).any()
            meets_criteria = meets_criteria and (positioned_domains >= min_sites)
            meets_criteria = meets_criteria and (domains_in_top_x >= min_sites_in_top)
        
        # Si le mot-clé répond aux critères, l'ajouter aux résultats
        if meets_criteria:
            # Créer une ligne avec les informations de base
            row = {'Mot clé': keyword}
            
            # Volume de recherche (prendre le maximum)
            volume = 0
            if 'volume' in group.columns:
                volume = group['volume'].max()
            
            row['Recherches mensuelles'] = int(volume)
            row['Nbre de NDD Positionnés'] = positioned_domains
            
            # Ajouter une colonne pour chaque domaine avec sa position
            for domain in domains:
                domain_data = group[group['domain'] == domain]
                if not domain_data.empty:
                    best_position = domain_data['position'].min()
                    if best_position < 100:  # Ne montrer que les positions < 100
                        row[f"Position_{domain}"] = int(best_position)
                    else:
                        row[f"Position_{domain}"] = None
                    
                    # Ajouter l'URL pour chaque domaine
                    if 'current_url' in domain_data.columns:
                        try:
                            best_idx = domain_data['position'].idxmin()
                            url = domain_data.loc[best_idx, 'current_url']
                            if pd.notna(url) and domain_data.loc[best_idx, 'position'] < 100:
                                row[f"URL_{domain}"] = url
                            else:
                                row[f"URL_{domain}"] = None
                        except:
                            row[f"URL_{domain}"] = None
                else:
                    row[f"Position_{domain}"] = None
                    row[f"URL_{domain}"] = None
            
            audit_results.append(row)
    
    status_text.text(f"Mots-clés correspondant aux critères: {len(audit_results)}/{total_keywords}")
    
    if not audit_results:
        return pd.DataFrame()
    
    # Créer le DataFrame final
    audit_df = pd.DataFrame(audit_results)
    
    # Trier par volume décroissant
    if 'Recherches mensuelles' in audit_df.columns:
        audit_df = audit_df.sort_values('Recherches mensuelles', ascending=False)
    
    return audit_df
# Traitement spécial pour l'onglet de compétition
    if sheet_name == "Compétition":
            # Réorganiser les colonnes pour regrouper toutes les positions ensemble, puis toutes les URLs ensemble
            base_cols = ['Mot clé', 'Recherches mensuelles', 'Nbre de NDD Positionnés']
            position_cols = []
            url_cols = []
            
            for col in df.columns:
                if col.startswith('Position_'):
                    domain = col.split('_', 1)[1]
                    # Extraire simplement le nom de domaine principal du fichier
                    clean_domain = domain.split('-')[0]  # Prend la première partie avant le tiret
                    # Renommer la colonne avec le nom de domaine propre
                    new_col_name = f"Position_{clean_domain}"
                    df = df.rename(columns={col: new_col_name})
                    position_cols.append(new_col_name)
                elif col.startswith('URL_'):
                    domain = col.split('_', 1)[1]
                    # Extraire simplement le nom de domaine principal du fichier
                    clean_domain = domain.split('-')[0]  # Prend la première partie avant le tiret
                    # Renommer la colonne avec le nom de domaine propre
                    new_col_name = f"URL_{clean_domain}"
                    df = df.rename(columns={col: new_col_name})
                    url_cols.append(new_col_name)
            
            # Trier les colonnes de positions et URLs par nom de domaine
            position_cols = sorted(position_cols)
            url_cols = sorted(url_cols)
            
            # Réordonner les colonnes : d'abord les colonnes de base, puis toutes les positions, puis toutes les URLs
            df = df[base_cols + position_cols + url_cols]
            
            # Écrire le DataFrame
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            st.dataframe(df)
            worksheet = writer.sheets[sheet_name]
            
            # Formater les en-têtes
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Format pour la colonne volume
            worksheet.set_column(1, 1, 20, volume_format)
            
            # Appliquer le code couleur pour les positions et formater les URLs
            for col_idx, col_name in enumerate(df.columns):
                if col_name.startswith('Position_'):
                    for row_idx in range(1, len(df) + 1):
                        try:
                            cell_value = df.iloc[row_idx-1][col_name]
                            if pd.notna(cell_value):
                                if 1 <= cell_value <= 3:
                                    worksheet.write(row_idx, col_idx, cell_value, pos_1_3_format)
                                elif 4 <= cell_value <= 10:
                                    worksheet.write(row_idx, col_idx, cell_value, pos_4_10_format)
                                elif 11 <= cell_value <= 20:
                                    worksheet.write(row_idx, col_idx, cell_value, pos_11_20_format)
                                else:
                                    worksheet.write(row_idx, col_idx, cell_value, pos_20plus_format)
                        except:
                            continue
                elif col_name.startswith('URL_'):
                    for row_idx in range(1, len(df) + 1):
                        try:
                            url = df.iloc[row_idx-1][col_name]
                            if pd.notna(url):
                                # Limiter les erreurs dues à la limite d'URLs
                                try:
                                    worksheet.write_url(row_idx, col_idx, url, string=url, cell_format=url_format)
                                except:
                                    worksheet.write(row_idx, col_idx, url)
                        except:
                            continue
            
            # Ajuster la largeur des colonnes
            worksheet.set_column(0, 0, 30)  # Mot-clé
            worksheet.set_column(1, 1, 20)  # Volume
            worksheet.set_column(2, 2, 15)  # Nombre de domaines
            
            # Ajuster la largeur pour les colonnes de positions et URLs
            for i, col_name in enumerate(df.columns[3:], start=3):
                if col_name.startswith('Position_'):
                    worksheet.set_column(i, i, 15)
                elif col_name.startswith('URL_'):
                    worksheet.set_column(i, i, 40)
            
            # Finitions
            worksheet.freeze_panes(1, 0)
            worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)
            
    else:
            # Traitement standard pour les autres onglets
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            worksheet = writer.sheets[sheet_name]
            
            # En-têtes
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Largeur des colonnes
            for idx, col in enumerate(df.columns):
                max_len = max(
                    df[col].astype(str).map(len).max() if len(df) > 0 else 10,
                    len(str(col))
                ) + 2
                worksheet.set_column(idx, idx, min(max_len, 50))
            
            # Filtres
            worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)
    
    writer.close()
    return output.getvalue()

def get_download_link(df_dict, filename="audit_semantique.xlsx"):
    """Crée un lien de téléchargement pour le fichier Excel."""
    val = to_excel(df_dict)
    b64 = base64.b64encode(val)
    return f'<a href="data:application/octet-stream;base64,{b64.decode()}" download="{filename}" class="download-button">Télécharger l\'audit sémantique</a>'
def main():
    st.title("Audit Sémantique SEO")
    
    # Style CSS pour le bouton de téléchargement
    st.markdown("""""
    <style>
    .download-button {
        display: inline-block;
        padding: 0.5em 1em;
        color: white;
        background-color: #4CAF50;
        text-align: center;
        text-decoration: none;
        font-size: 16px;
        border-radius: 4px;
        margin: 10px 0;
    }
    .stProgress > div > div > div > div {
        background-color: #4CAF50;
    }
    </style>
    """, unsafe_allow_html=True)
    
    # Interface simplifiée
    st.subheader("Importer les fichiers d'export Ahrefs")
    uploaded_files = st.file_uploader("Déposez vos fichiers ici (un fichier par domaine)", 
                                     accept_multiple_files=True, 
                                     type=['csv', 'xlsx', 'xls'])
    
    if not uploaded_files:
        st.info("Veuillez importer des fichiers d'export Ahrefs (CSV ou Excel)")
        st.markdown("""
        **Conseils pour de meilleurs résultats:**
        - Nommez vos fichiers avec le nom du domaine (ex: example.com.xlsx)
        - Assurez-vous que vos exports contiennent les colonnes Keyword, Volume, Current position et Current URL
        """)
        return
    
    # Afficher les fichiers importés sans détails techniques
    file_names = [file.name for file in uploaded_files]
    st.write(f"Fichiers importés: {len(file_names)}")
    
    # Configuration simplifiée
    col1, col2 = st.columns(2)
    
    with col1:
        # Mapping des colonnes simplifié
        st.subheader("Configuration des colonnes")
        column_mapping = {
            'keyword': st.text_input("Colonne Mot-clé:", "Keyword"),
            'volume': st.text_input("Colonne Volume:", "Volume"),
            'position': st.text_input("Colonne Position:", "Current position"),
            'current_url': st.text_input("Colonne URL:", "Current URL")
        }
    
    with col2:
        # Filtres simplifiés et améliorés
        st.subheader("Filtres")
        filter_type = st.selectbox("Type de filtre:", [
            "Au moins 1 site positionné dans le top 10",
            "Au moins 1 site positionné dans le top 20",
            "Au moins 1 site positionné dans le top 30",
            "Au moins 2 sites positionnés, dont 1 top 10",
            "Au moins 2 sites positionnés, dont 1 top 20",
            "Au moins 2 sites positionnés, dont 1 top 30"
        ])
        
        # Déterminer automatiquement la valeur de top_positions en fonction du filtre choisi
        if "top 10" in filter_type:
            default_top_pos = 10
        elif "top 20" in filter_type:
            default_top_pos = 20
        else:  # "top 30"
            default_top_pos = 30
        
        # Déterminer automatiquement le nombre minimum de sites en fonction du filtre
        if filter_type.startswith("Au moins 2 sites"):
            default_min_sites = 2
        else:
            default_min_sites = 1
        
        min_sites = st.number_input("Nombre minimum de sites se positionnant sur le mot-clé:", 
                                   min_value=1, value=default_min_sites)
        top_positions = st.number_input("Nombre minimum de sites se positionnants dans les X premières positions:", 
                                      min_value=1, value=default_top_pos)
        min_sites_in_top = st.number_input("Nombre minimum de sites dans les X premières positions:", 
                                         min_value=1, value=1)
    
    # Filtres
    filters = {
        'filter_type': filter_type,
        'min_sites': min_sites,
        'top_positions': top_positions,
        'min_sites_in_top': min_sites_in_top
    }
    
    # Bouton d'analyse
    if st.button("Générer l'audit sémantique"):
        if uploaded_files:
            # Traitement silencieux
            with st.spinner("Analyse en cours..."):
                # Traiter les fichiers
                combined_data, source_data = process_files(uploaded_files, column_mapping)
                
                if combined_data is not None and not combined_data.empty:
                    # Créer l'audit
                    audit_df = create_competition_audit(combined_data, filters)
                    
                    if audit_df.empty:
                        st.warning("Aucun résultat ne correspond aux critères sélectionnés.")
                    else:
                        # Créer le fichier Excel
                        df_dict = {
                            "Compétition": audit_df,
                            "Sources": source_data
                        }
                        
                        # Message de succès et lien de téléchargement
                        st.success(f"✅ Audit sémantique généré avec succès ({len(audit_df)} mots-clés)")
                        st.markdown(get_download_link(df_dict), unsafe_allow_html=True)
                else:
                    st.error("Erreur lors du traitement des fichiers. Vérifiez que vos fichiers sont au format attendu.")
def to_excel(df_dict):
    """Génère un fichier Excel formaté avec des colonnes séparées pour les URLs."""
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    
    workbook = writer.book
    
    # Formats
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#D7E4BC',
        'border': 1
    })
    
    # Formats couleurs pour les positions
    pos_1_3_format = workbook.add_format({'bg_color': '#90EE90'})  # Vert clair
    pos_4_10_format = workbook.add_format({'bg_color': '#FFFF99'})  # Jaune
    pos_11_20_format = workbook.add_format({'bg_color': '#FFB6C1'})  # Rose
    pos_20plus_format = workbook.add_format({'bg_color': '#FF6666'})  # Rouge
    
    # Format pour les URLs
    url_format = workbook.add_format({
        'font_color': 'blue',
        'underline': True
    })
    
    # Format numérique pour les volumes
    volume_format = workbook.add_format({'num_format': '#,##0'})
    
    # Traiter chaque onglet
    for sheet_name, df in df_dict.items():
        if df.empty:
            continue
            
        # Traitement spécial pour l'onglet de compétition
        if sheet_name == "Compétition":
            # Réorganiser les colonnes pour regrouper toutes les positions ensemble, puis toutes les URLs ensemble
            base_cols = ['Mot clé', 'Recherches mensuelles', 'Nbre de NDD Positionnés']
            position_cols = []
            url_cols = []
            
            for col in df.columns:
                if col.startswith('Position_'):
                    domain = col.split('_', 1)[1]
                    # Extraire simplement le nom de domaine principal du fichier
                    clean_domain = domain.split('-')[0] if '-' in domain else domain  # Prend la première partie avant le tiret
                    # Renommer la colonne avec le nom de domaine propre
                    new_col_name = f"Position_{clean_domain}"
                    df = df.rename(columns={col: new_col_name})
                    position_cols.append(new_col_name)
                elif col.startswith('URL_'):
                    domain = col.split('_', 1)[1]
                    # Extraire simplement le nom de domaine principal du fichier
                    clean_domain = domain.split('-')[0] if '-' in domain else domain  # Prend la première partie avant le tiret
                    # Renommer la colonne avec le nom de domaine propre
                    new_col_name = f"URL_{clean_domain}"
                    df = df.rename(columns={col: new_col_name})
                    url_cols.append(new_col_name)
            
            # Trier les colonnes de positions et URLs par nom de domaine
            position_cols = sorted(position_cols)
            url_cols = sorted(url_cols)
            
            # Réordonner les colonnes : d'abord les colonnes de base, puis toutes les positions, puis toutes les URLs
            df = df[base_cols + position_cols + url_cols]
            
            # Écrire le DataFrame
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            st.dataframe(df)
            worksheet = writer.sheets[sheet_name]
            
            # Formater les en-têtes
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Format pour la colonne volume
            worksheet.set_column(1, 1, 20, volume_format)
            
            # Appliquer le code couleur pour les positions et formater les URLs
            for col_idx, col_name in enumerate(df.columns):
                if col_name.startswith('Position_'):
                    for row_idx in range(1, len(df) + 1):
                        try:
                            cell_value = df.iloc[row_idx-1][col_name]
                            if pd.notna(cell_value):
                                if 1 <= cell_value <= 3:
                                    worksheet.write(row_idx, col_idx, cell_value, pos_1_3_format)
                                elif 4 <= cell_value <= 10:
                                    worksheet.write(row_idx, col_idx, cell_value, pos_4_10_format)
                                elif 11 <= cell_value <= 20:
                                    worksheet.write(row_idx, col_idx, cell_value, pos_11_20_format)
                                else:
                                    worksheet.write(row_idx, col_idx, cell_value, pos_20plus_format)
                        except:
                            continue
                elif col_name.startswith('URL_'):
                    for row_idx in range(1, len(df) + 1):
                        try:
                            url = df.iloc[row_idx-1][col_name]
                            if pd.notna(url):
                                # Limiter les erreurs dues à la limite d'URLs
                                try:
                                    worksheet.write_url(row_idx, col_idx, url, string=url, cell_format=url_format)
                                except:
                                    worksheet.write(row_idx, col_idx, url)
                        except:
                            continue
            
            # Ajuster la largeur des colonnes
            worksheet.set_column(0, 0, 30)  # Mot-clé
            worksheet.set_column(1, 1, 20)  # Volume
            worksheet.set_column(2, 2, 15)  # Nombre de domaines
            
            # Ajuster la largeur pour les colonnes de positions et URLs
            for i, col_name in enumerate(df.columns[3:], start=3):
                if col_name.startswith('Position_'):
                    worksheet.set_column(i, i, 15)
                elif col_name.startswith('URL_'):
                    worksheet.set_column(i, i, 40)
            
            # Finitions
            worksheet.freeze_panes(1, 0)
            worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)
            
        else:
            # Traitement standard pour les autres onglets
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            worksheet = writer.sheets[sheet_name]
            
            # En-têtes
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Largeur des colonnes
            for idx, col in enumerate(df.columns):
                max_len = max(
                    df[col].astype(str).map(len).max() if len(df) > 0 else 10,
                    len(str(col))
                ) + 2
                worksheet.set_column(idx, idx, min(max_len, 50))
            
            # Filtres
            worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)
    
    writer.close()
    return output.getvalue()

def get_download_link(df_dict, filename="audit_semantique.xlsx"):
    """Crée un lien de téléchargement pour le fichier Excel."""
    val = to_excel(df_dict)
    b64 = base64.b64encode(val)
    return f'<a href="data:application/octet-stream;base64,{b64.decode()}" download="{filename}" class="download-button">Télécharger l\'audit sémantique</a>'

if __name__ == "__main__":
    main()