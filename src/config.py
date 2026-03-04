import json

CONFIG_PATH = 'config/settings.json'

def carregar_config():
    with open(CONFIG_PATH, 'r', encoding='utf-8') as arquivo:
        return json.load(arquivo)
    
if __name__ == "__main__":
    config = carregar_config()
    print(config)