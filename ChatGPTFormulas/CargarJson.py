import json
def cargar_configuracion(file_path):
    with open(file_path, 'r') as file:
        config = json.load(file)
    return config
