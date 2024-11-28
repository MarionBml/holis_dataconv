# import openpyxl
import pandas
import sqlite3

# Noms des fichiers utilisés
excel_file = "BI_2.02__06_CatImpacts_Details.xlsx" 
db_file = "ma_base.db"
table_name = "CatImpacts_table" 

# Chargement et transformation du dataframe
cat_impacts = pd.read_excel(excel_file, sheet_name=0, index_col=1) # pandas détecte automatiquement les types de données 
cat_impacts = cat_impacts.T
cat_impacts.columns = cat_impacts.columns.str.strip()
cat_impacts.set_index('UUID', inplace = True)
cat_impacts.rename(columns={cat_impacts.columns[1]: "French Name" }, inplace = True)
cat_impacts.drop(cat_impacts.index[0], inplace = True)
cat_impacts.drop(cat_impacts.columns[cat_impacts.nunique() == 1], axis=1, inplace=True)
cat_impacts['Dataset format'] = 'ILCD format'

# Connexion à SGBD (ici SQLite)
conn = sqlite3.connect(db_file) 

# Ecrit les infos stockées dans un DataFrame vers une base SQL
cat_impacts.to_sql(table_name, conn, if_exists='replace', index=False) # if_exists='replace' remplace la table si elle existe déjà, index=False évite d'ajouter un index 

# Imprime une phrase après l'écriture dans la base SQL
print(f"Feuille Excel convertie en table '{table_name}' dans la base de données '{db_file}'") 

# Fermeture de la connexion
conn.close() 
