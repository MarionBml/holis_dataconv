import pandas as pd
import sqlite3

# Noms des fichiers utilisés
excel_file = "CatImpacts_Details.xlsx" #Fichier Excel
db_file = "CatImpacts_sql.db"                  
table_name = "CatImpacts_table" 

# Chargement et transformation du dataframe
cat_impacts = pd.read_excel(excel_file, sheet_name=0, index_col=1) # On crée un dataframe à partir de l'excel
cat_impacts = cat_impacts.T                                          # On transpose pour avoir les caractéristiques en colonnes
cat_impacts.columns = cat_impacts.columns.str.strip()               # On enlève les espaces inutiles dans les noms de colonnes
cat_impacts.set_index('UUID', inplace = True)                          #La colonne UUID devient la colonne d'index 
cat_impacts.rename(columns={cat_impacts.columns[1]: "French Name" }, inplace = True) # On nomme "French name" la colonne vide
cat_impacts.drop(cat_impacts.index[0], inplace = True)                                  # On enlève les colonnes inutiles
cat_impacts.drop(cat_impacts.columns[cat_impacts.nunique() == 1], axis=1, inplace=True)
cat_impacts['Dataset format'] = 'ILCD format'                               

# Connexion à SGBD, ici SQLite
conn = sqlite3.connect(db_file) 

# Ecrit les infos stockées dans le DataFrame vers une base SQL
cat_impacts.to_sql(table_name, conn, if_exists='replace', index=False) # if_exists='replace' remplace la table si elle existe déjà, index=False évite d'ajouter un index 

# Imprime une phrase après l'écriture dans la base SQL
print(f"Feuille Excel convertie en table '{table_name}' dans la base de données '{db_file}'") 

# Fermeture de la connexion
conn.close() 
