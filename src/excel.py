import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

import pandas as pd
from src.config import carregar_config

def ler_planilha(mes, ano):
    config = carregar_config()
    caminho = config['caminhos']['planilha_base']
    nome_aba = f"Cobrança {mes:02d}.{ano}"

    df = pd.read_excel(caminho, sheet_name=nome_aba)
    df = df[df['OS'].notna()]
    df = df[df['OS'] != 'OCIOSO'] 
    
    oss = df['OS'].unique()
    return df, oss

if __name__ == "__main__":
    from datetime import datetime
    hoje = datetime.now()
    df, oss = ler_planilha(hoje.month - 1, hoje.year)
    print(f"OSs encontradas: {oss}")
    print(f"Total de linhas: {len(df)}")