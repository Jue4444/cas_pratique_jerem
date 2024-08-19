import pandas as pd

def parse_data_from_excel(file_path):
    df = pd.read_excel(file_path)
    
    # Suppression des lignes vides ou presque vides
    df = df.dropna(how='all')
    
    # Extraire les noms, prénoms et emails
    names = df.iloc[::10, 1].tolist()  # Les noms/prénoms sont dans la 2ème colonne
    emails = df.iloc[3::10, 1].apply(lambda x: x.split('/')[-1].strip()).tolist()  # Les emails sont dans la 2ème colonne aussi, mais 3 lignes après les noms

    # Rassembler les données extraites
    extracted_data = pd.DataFrame({
        "Nom et Prénom": names,
        "Email": emails
    })

    return extracted_data

def save_to_excel(data, output_file_path):
    data.to_excel(output_file_path, index=False)

# Chemin du fichier d'entrée
input_file_path = "01.xls"
# Chemin du fichier de sortie
output_file_path = "resultat_final.xlsx"

# Extraire les données du fichier Excel d'entrée
data = parse_data_from_excel(input_file_path)
# Sauvegarder les données extraites dans un nouveau fichier Excel
save_to_excel(data, output_file_path)

print(f"Le fichier {output_file_path} a été généré avec succès.")
