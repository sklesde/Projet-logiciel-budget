from os import listdir, path, remove, makedirs
from pandas import read_csv, read_excel, to_datetime, to_numeric, ExcelWriter
import pandas as pd
from datetime import date
import shutil
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font


#%%
   
def ouverture_csv(date):
    try:
        # Recherche du premier fichier CSV dans le répertoire courant
        csv_files = [f for f in listdir() if f.endswith(".csv")]
        
        if len(csv_files) == 0:
            raise FileNotFoundError("Aucun fichier CSV trouvé dans le dossier.")
            """
            Par la suite, si aucun dossier n'est trouvé on renvoiera qu'aucun dossier n'a été trouve
            mais que la catégorisation de Budget excel à été mise à jours.
            """
        elif len(csv_files) > 1:
            raise ValueError("Plusieurs fichiers CSV sont présents dans le dossier.")
        
        path_data= csv_files[0]
        data = read_csv(path_data, encoding="utf-8", sep=";")
        data.to_excel(f".\Stockage CSV Banque\{date}.xlsx")
        print(f"Le CSV rentré à été stocké sous le nom de {date} dans le répertoire Stockage CSV Banque.")
        
        L=(data, path_data)
        return L
    
    except FileNotFoundError as e:
        print(str(e))
        return (0,0)
    except ValueError as e:
        print(str(e))
        return (0,0)
    except Exception as e:
        print(f"Une erreur inattendue s'est produite : {str(e)}")
        return (0,0)

def verifier_dossiers():
    dossier_precedentes_versions = "Précédentes versions"
    dossier_stockage_csv_banque = "Stockage CSV Banque"
    
    # Vérification de l'existence des dossiers
    existe_precedentes_versions = path.exists(dossier_precedentes_versions)
    existe_stockage_csv_banque = path.exists(dossier_stockage_csv_banque)
    
    # Si le dossier "Précédentes versions" n'existe pas, on l'ajoute au compteur et on le crée
    if not existe_precedentes_versions:
        makedirs(dossier_precedentes_versions)
        print('Le dossier "Précédentes versions" à été créé.')
    
    # Si le dossier "Stockage CSV Banque" n'existe pas, on l'ajoute au compteur et on le crée
    if not existe_stockage_csv_banque:
        makedirs(dossier_stockage_csv_banque)
        print('Le dossier "Stockage CSV Banque" à été créé.')
        
    return

def convertisseur_en_chiffre(df, columns):
    """
    Convertit les colonnes spécifiées d'un DataFrame en type float.
    Les valeurs non convertibles sont remplacées par NaN.
    
    :param df: DataFrame pandas
    :param columns: Liste des noms de colonnes à convertir en float
    :return: DataFrame avec les colonnes converties
    """
    for column in columns:
        if column in df.columns:
            # Nettoyer les valeurs en remplaçant les virgules par des points
            # et en supprimant les autres caractères non numériques (comme les symboles)
            df[column] = df[column].str.replace(',', '.', regex=False)  # Remplacer la virgule par le point
            
            # Supprimer les caractères non numériques, à l'exception du point et du signe moins
            df[column] = df[column].str.replace(r'[^\d.-]', '', regex=True)
            
            # Conversion des colonnes en float, avec gestion des erreurs
            df[column] = to_numeric(df[column], errors='coerce')
    
    return df



def generer_id_unique(df, date_col="Date operation", libelle_col="Libelle operation"):
    """
    Génère un ID unique pour chaque transaction en fonction du jour et du libellé d'opération.
    Si plusieurs transactions sont identiques sur un même jour, elles sont numérotées (A1, A2, B1, etc.).
    
    :param df: DataFrame contenant les transactions
    :param date_col: Nom de la colonne contenant la date
    :param libelle_col: Nom de la colonne contenant le libellé d'opération
    :return: DataFrame avec une colonne 'ID' unique
    """
    df = df.copy()
    
    # Assurer que la colonne date est bien au format datetime
    df[date_col] = to_datetime(df[date_col], format='%d/%m/%Y', errors='coerce')
    
    # Convertir la date en format string pour l'utiliser dans l'ID (sans l'heure)
    df["date_str"] = df[date_col].dt.strftime("%Y%m%d")  # Seulement la date sans l'heure


    # Grouper par date et libellé d'opération et générer des indices uniques
    df["ID_count"] = df.groupby(["date_str", libelle_col]).cumcount() + 1

    # Création de l'ID unique
    df['ID'] = df.apply(lambda x: f"{x['date_str']}_{x[libelle_col][:10].upper()}_{x['ID_count']}", axis=1)

    # Vérification des doublons et correction
    unique_ids = set()
    for i in df.index:
        id_val = df.at[i, "ID"]
        while id_val in unique_ids:
            df.at[i, "ID_count"] += 1
            id_val = f"{df.at[i, 'date_str']}_{df.at[i, libelle_col][:10].upper()}_{df.at[i, 'ID_count']}"
        unique_ids.add(id_val)
        df.at[i, "ID"] = id_val

    # Suppression des colonnes temporaires
    df = df.drop(columns=["date_str", "ID_count"])

    return df

def creation_data_cp(data, budget_mensuel_donnees):
    try :
        #Je crée une copie de data pour manipuler les données
        colonnes_a_garder = ["Date operation", "Libelle simplifie", "Libelle operation", "Categorie", "Sous categorie", "Debit", "Credit"]

        data_cp = data[colonnes_a_garder].copy() 
        colonnes=["Debit", "Credit"]
        data_cp = convertisseur_en_chiffre(data_cp, colonnes)

        #Je crée la colonne ID et Classification avec des None à l'intérieur
        data_cp['ID']=None
        data_cp['Classification']=None

        #Je crée les ID uniques
        data_cp = generer_id_unique(data_cp)
        
        
        
        return data_cp
        print('test')
    except:
        print("Pas de fichier CSV détecté, mise à jours des catégories.")
    return 0


def verifier_et_organiser_colonnes(feuille_excel: pd.DataFrame, colonnes_attendues: list) -> pd.DataFrame:
    # Si le DataFrame est vide, créer un DataFrame avec les colonnes spécifiées
    if feuille_excel.empty:
        feuille_excel = pd.DataFrame(columns=colonnes_attendues)
    
    # Réorganiser les colonnes pour correspondre à colonnes_attendues, peu importe l'ordre
    feuille_excel = feuille_excel.reindex(columns=colonnes_attendues)
    
    return feuille_excel



def verification_et_fusion(df1, df2, id_col="ID"):
    # Vérification si df2 est un entier
    if isinstance(df2, int):
        return df1
    
    # Vérification que les ID du df2 ne sont pas déjà dans df1
    ids_df1 = set(df1[id_col])
    df2_filtered = df2[~df2[id_col].isin(ids_df1)]
    
    # Filtrage des colonnes vides ou remplies uniquement de NaN dans df2_filtered
    df2_filtered = df2_filtered.dropna(axis=1, how='all')
    
    # Fusion des deux DataFrames
    df_merged = pd.concat([df1, df2_filtered], ignore_index=True)

    return df_merged

def mettre_a_jour_classification(budget_mensuel_donnees, budget_mensuel_categories):
    # Filtrer les lignes où la classification n'est pas NaN ou vide
    filtered_data = budget_mensuel_categories.loc[budget_mensuel_categories['Classification'].notna() & (budget_mensuel_categories['Classification'] != '')]

    for i in range(len(filtered_data)):
        # Extraire la ligne à copier en tant que DataFrame
        ligne_a_copier = filtered_data.iloc[[i]]
        
        # Trouver l'index de la ligne dans budget_mensuel_donnees où l'ID correspond
        index_ligne_donnes = budget_mensuel_donnees[budget_mensuel_donnees['ID'] == filtered_data.loc[i, 'ID']].index
        
        # Vérifier si un index correspondant est trouvé et mettre à jour la ligne
        if len(index_ligne_donnes) > 0:
            # Pour éviter les problèmes de type, on peut d'abord s'assurer que les types correspondent
            for col in ligne_a_copier.columns:
                # Exemple de conversion explicite : convertir les colonnes numériques en float64
                if budget_mensuel_donnees[col].dtype == 'float64' and pd.api.types.is_numeric_dtype(ligne_a_copier[col]):
                    ligne_a_copier[col] = ligne_a_copier[col].astype('float64')
                
                # Si la colonne attend une chaîne (object), assure-toi qu'elle soit bien une chaîne
                if budget_mensuel_donnees[col].dtype == 'object' and pd.api.types.is_string_dtype(ligne_a_copier[col]):
                    ligne_a_copier[col] = ligne_a_copier[col].astype('object')

            # Mettre à jour la ligne dans budget_mensuel_donnees
            budget_mensuel_donnees.loc[index_ligne_donnes[0]] = ligne_a_copier.iloc[0]

    # Refiltrer budget_mensuel_categories avec les lignes où 'Classification' est NaN ou vide
    budget_mensuel_categories = budget_mensuel_donnees.loc[budget_mensuel_donnees['Classification'].isna() | (budget_mensuel_donnees['Classification'] == '')]
    
    return budget_mensuel_donnees, budget_mensuel_categories

def tri_par_semaine(df):
    # S'assurer qu'on travaille sur une copie explicite pour éviter les avertissements
    df = df.copy()

    # Créer la colonne 'Début de Semaine' avec le premier jour de chaque semaine
    df['Début de Semaine'] = df['Date operation'].dt.to_period('W').dt.start_time.dt.date
    
    # Créer la colonne 'Fin de Semaine' avec le dernier jour de chaque semaine
    df['Fin de Semaine'] = (df['Début de Semaine'] + pd.Timedelta(days=6))

    # Créer la colonne 'Semaine' avec la plage de dates "Début de Semaine - Fin de Semaine"
    df['Semaine'] = df['Début de Semaine'].astype(str) + ' - ' + df['Fin de Semaine'].astype(str)
    
    # Trier le DataFrame par la colonne 'Semaine'
    df = df.sort_values(by='Semaine')
    
    # Retourner le DataFrame avec la colonne 'Semaine' ajoutée
    return df


def calcul_et_tri(df):
    # S'assurer qu'on travaille sur une copie explicite pour éviter les avertissements
    df = df.copy()

    # Créer la colonne 'Début de Semaine' avec le premier jour de chaque semaine
    df['Début de Semaine'] = df['Date operation'].dt.to_period('W').dt.start_time.dt.date
    
    # Créer la colonne 'Fin de Semaine' avec le dernier jour de chaque semaine
    df['Fin de Semaine'] = (df['Début de Semaine'] + pd.Timedelta(days=6))

    # Créer la colonne 'Semaine' avec la plage de dates "Début de Semaine - Fin de Semaine"
    df['Semaine'] = df['Début de Semaine'].astype(str) + ' - ' + df['Fin de Semaine'].astype(str)
    
    # Filtrer les lignes où 'Classification' appartient à la liste spécifiée
    classifications_valide = [
        "Courses", "Snacks", "Restaurants", "Sport", "Vêtements/Coiffure", 
        "Loisirs", "Divers", "Commande Internet", "Transports", "Autre 1", "Autre 2"
    ]
    df_filtre = df[df['Classification'].isin(classifications_valide)]
    
    # Vérifier si les colonnes 'Debit' et 'Credit' existent dans le DataFrame
    if 'Debit' not in df.columns or 'Credit' not in df.columns:
        raise KeyError("Les colonnes 'Debit' et 'Credit' ne sont pas présentes dans le DataFrame.")
    
    # Calculer la somme des colonnes "Debit" et "Credit" pour chaque combinaison de 'Semaine' et 'Classification'
    df_somme = df_filtre.groupby(['Semaine', 'Classification'])[['Debit', 'Credit']].sum().reset_index()
    
    # Ajouter une colonne 'Total' qui est la somme de 'Debit' et 'Credit'
    df_somme['Total'] = df_somme['Debit'] + df_somme['Credit']
    
    # Garder uniquement les colonnes 'Semaine', 'Classification' et 'Total'
    df_somme = df_somme[['Semaine', 'Classification', 'Total']]
    
    return df_somme

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side
import pandas as pd

def envoie_donnees(df, file_path):
    nom_feuille = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
    mois_dict = {i+1: nom_feuille[i] for i in range(12)}
    classifications_depenses = [
        "Courses", "Snacks", "Restaurants", "Sport", "Vêtements/Coiffure", 
        "Loisirs", "Divers", "Commande Internet", "Transports", "Autre 1", "Autre 2"
    ]
    
    try:
        wb = load_workbook(file_path)
    except FileNotFoundError:
        raise FileNotFoundError(f"Le fichier {file_path} n'a pas été trouvé.")
    
    # Définir les bordures
    border_thick_left_right = Border(
        left=Side(style="thick"),
        right=Side(style="thick")
    )
    border_thick_all_sides = Border(
        top=Side(style="thick"),
        left=Side(style="thick"),
        bottom=Side(style="thick"),
        right=Side(style="thick")
    )
    
    for ws in wb.worksheets:
        # Défusionner les cellules dans la zone spécifiée
        for row in range(12, ws.max_row + 1):
            for col in [9, 10]:  # Colonnes I et J
                cell = ws.cell(row=row, column=col)
                
                # Vérifier si la cellule fait partie d'une plage fusionnée
                for merged_range in ws.merged_cells.ranges:
                    if cell.coordinate in merged_range:
                        # Défusionner la cellule
                        ws.unmerge_cells(str(merged_range))
                        break  # Une fois la cellule défusionnée, on peut sortir de la boucle
                
                # Maintenant, on peut modifier la valeur de la cellule
                cell.value = None
                cell.fill = PatternFill(fill_type=None)
                
                # Appliquer les bordures extérieures
                if col == 9:  # Colonne I (Gauche)
                    cell.border = border_thick_left_right
                if col == 10:  # Colonne J (Droite)
                    cell.border = border_thick_left_right
    
    semaine_data = {}
    
    for index, row in df.iterrows():
        semaine = row["Semaine"]
        classification = row["Classification"]
        total = row["Total"]
        
        if classification not in classifications_depenses:
            continue
        
        if semaine not in semaine_data:
            semaine_data[semaine] = []
        semaine_data[semaine].append((classification, total))
    
    mois_totaux = {}
    
    for semaine, data in semaine_data.items():
        start_date, end_date = semaine.split(" - ")
        start_date = pd.to_datetime(start_date)
        end_date = pd.to_datetime(end_date)
        
        jours_par_mois = {i: 0 for i in range(1, 13)}
        for single_date in pd.date_range(start=start_date, end=end_date):
            mois = single_date.month
            jours_par_mois[mois] += 1
        
        mois_max = max(jours_par_mois, key=jours_par_mois.get)
        mois_nom = mois_dict[mois_max]
        
        if mois_nom not in wb.sheetnames:
            raise ValueError(f"La feuille {mois_nom} n'existe pas dans le fichier.")
        
        ws = wb[mois_nom]
        row_to_insert = 12
        while ws.cell(row=row_to_insert, column=9).value is not None:
            row_to_insert += 1
        
        week_cell = ws.cell(row=row_to_insert, column=9, value=f"{semaine}")
        week_cell.fill = PatternFill(start_color="FFCC99", end_color="FFCC99", fill_type="solid")
        week_cell.border = border_thick_left_right  # Bordure épaisse à gauche et droite pour la semaine
        
        row_to_insert += 1
        total_depenses_semaine = 0
        
        for classification, total in data:
            ws.cell(row=row_to_insert, column=9, value=classification).font = Font(bold=False)
            ws.cell(row=row_to_insert, column=10, value=total).font = Font(bold=False)
            ws.cell(row=row_to_insert, column=9).border = border_thick_left_right  # Bordure épaisse à gauche pour la colonne I
            ws.cell(row=row_to_insert, column=10).border = border_thick_left_right  # Bordure épaisse à droite pour la colonne J
            total_depenses_semaine += total
            row_to_insert += 1
        
        ws.cell(row=row_to_insert, column=9, value="TOTAL DEPENSES SEMAINE").font = Font(bold=True)
        ws.cell(row=row_to_insert, column=10, value=total_depenses_semaine).font = Font(bold=True)
        ws.cell(row=row_to_insert, column=9).border = border_thick_left_right  # Bordure épaisse à gauche pour le total
        ws.cell(row=row_to_insert, column=10).border = border_thick_left_right  # Bordure épaisse à droite pour le total
        row_to_insert += 1
        
        if mois_nom not in mois_totaux:
            mois_totaux[mois_nom] = 0
        mois_totaux[mois_nom] += total_depenses_semaine
    
    for mois_nom, total_mois in mois_totaux.items():
        ws = wb[mois_nom]
        row_to_insert = 12
        while ws.cell(row=row_to_insert, column=9).value is not None:
            row_to_insert += 1
        
        # Bordure épaisse à gauche et à droite pour "TOTAL DEPENSES MOIS"
        ws.cell(row=row_to_insert, column=9, value="TOTAL DEPENSES MOIS").font = Font(bold=True)
        ws.cell(row=row_to_insert, column=10, value=total_mois).font = Font(bold=True)
        ws.cell(row=row_to_insert, column=9).border = border_thick_left_right  # Bordure épaisse à gauche
        ws.cell(row=row_to_insert, column=10).border = border_thick_left_right  # Bordure épaisse à droite
        
        # Bordure épaisse sur les 4 côtés pour "TOTAL DEPENSES MOIS"
        ws.cell(row=row_to_insert, column=9).border = border_thick_all_sides
        ws.cell(row=row_to_insert, column=10).border = border_thick_all_sides
        
        row_to_insert += 1  # Appliquer les bordures jusqu'à "TOTAL DEPENSES MOIS"
        
        # Supprimer les bordures après "TOTAL DEPENSES MOIS"
        for row in range(row_to_insert, ws.max_row + 1):
            for col in [9, 10]:  # Colonnes I et J
                cell = ws.cell(row=row, column=col)
                cell.border = Border()  # Supprimer la bordure en la réinitialisant
                
    wb.save(file_path)




def enregistrement(data_cp, path_data, budget_mensuel_categories, budget_mensuel_donnees, file_path):
    # Vérifier si 'data_cp' est un entier
    if not isinstance(data_cp, int):

        remove(f".\{path_data}")
        
    with ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        # Réécrire uniquement la feuille 'Categories' modifiée
        budget_mensuel_categories.to_excel(writer, sheet_name='Categories', index=False)
    with ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        # Réécrire uniquement la feuille 'Donnees' modifiée
        budget_mensuel_donnees.to_excel(writer, sheet_name='Donnees', index=False)   

#%%
#On récupère la date du jour
date = date.today()
#Ouverture des nouvelles données à fusionner
data, path_data = ouverture_csv(date)
"""
A vérifier si je laisse le fichier excel Budget Mensuel dans le même dossier que l'exécutable et le reste.
"""
#On vérifier que les dossier existent
verifier_dossiers()

file_path="Budget Mensuel.xlsx"
data2 = read_excel(file_path)

destination_path = f'.\\Précédentes versions\\Budget Mensuel - Version du {date}.xlsx'

# Copier et renommer le fichier
shutil.copy(file_path, destination_path)

budget_mensuel_categories = read_excel(file_path, sheet_name="Categories")
budget_mensuel_donnees = read_excel(file_path, sheet_name="Donnees")

colonnes_attendues = ["Date operation", "Libelle simplifie", "Libelle operation", "Categorie", "Sous categorie", "Debit", "Credit", "ID", "Classification"]


budget_mensuel_donnees = verifier_et_organiser_colonnes(budget_mensuel_donnees, colonnes_attendues)
budget_mensuel_categories = verifier_et_organiser_colonnes(budget_mensuel_categories, colonnes_attendues)


data_cp = creation_data_cp(data, budget_mensuel_donnees)

budget_mensuel_donnees =verification_et_fusion(budget_mensuel_donnees, data_cp)

budget_mensuel_donnees, budget_mensuel_categories = mettre_a_jour_classification(budget_mensuel_donnees, budget_mensuel_categories)


data_calcul = budget_mensuel_donnees.loc[budget_mensuel_donnees['Classification'].notna() & (budget_mensuel_donnees['Classification'] != '')]

data_calcul = tri_par_semaine(data_calcul)

data_somme_semaines = calcul_et_tri(data_calcul)

envoie_donnees(data_somme_semaines, file_path)

enregistrement(data_cp, path_data, budget_mensuel_categories, budget_mensuel_donnees, file_path)

print("Excel mis a jour.")  
print("Appuyez sur une touche pour fermer...")










