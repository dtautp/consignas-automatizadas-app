import os

def create_folder_and_subfolder(file_name):
    # Divide el nombre del archivo para obtener el nombre de la subcarpeta
    subfolder_name = file_name.split("_")[1]
    
    # Define las rutas para la carpeta y subcarpeta
    folder_path = 'output'
    subfolder_path = f'output/{subfolder_name}'
    
    # Crea la carpeta principal si no existe
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    
    # Crea la subcarpeta si no existe
    if not os.path.exists(subfolder_path):
        os.makedirs(subfolder_path)

    return subfolder_path