import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

import comtypes.client
from pathlib import Path
from src.database import get_conn


def converter_para_pdf(caminho_xlsx):
    caminho_xlsx = Path(caminho_xlsx)
    caminho_pdf = caminho_xlsx.with_suffix('.pdf')
    
    try:
        excel = comtypes.client.CreateObject('Excel.Application')
        excel.Visible = False
        
        wb = excel.Workbooks.Open(str(caminho_xlsx.absolute()))
        wb.ExportAsFixedFormat(0, str(caminho_pdf.absolute()))
        wb.Close(False)
        excel.Quit()
        
        print(f"PDF gerado: {caminho_pdf.name}")
        return str(caminho_pdf)
        
    except Exception as e:
        print(f"Erro ao converter {caminho_xlsx.name}: {e}")
        return None

def converter_todos_bms():
    conn = get_conn()
    cursor = conn.cursor()
    
    cursor.execute('''
        SELECT id, caminho_xlsx FROM boletins WHERE status = ?
    ''', ('criado',))
    
    bms = cursor.fetchall()
    conn.close()
    
    if not bms:
        print("Nenhum BM pra converter!")
        return
    
    for bm_id, caminho_xlsx in bms:
        caminho_pdf = converter_para_pdf(caminho_xlsx)
        
        if caminho_pdf:
            conn = get_conn()
            cursor = conn.cursor()
            cursor.execute('''
                UPDATE boletins SET caminho_pdf = ? WHERE id = ?
            ''', (caminho_pdf, bm_id))
            conn.commit()
            conn.close()

def converter_todas_faturas():
    conn = get_conn()
    cursor = conn.cursor()
    
    cursor.execute('''
        SELECT id, caminho_xlsx FROM faturas WHERE status IN ('criada', 'pronta')
    ''')
    
    faturas = cursor.fetchall()
    conn.close()
    
    if not faturas:
        print("Nenhuma fatura pra converter!")
        return
    
    for fat_id, caminho_xlsx in faturas:
        caminho_pdf = converter_para_pdf(caminho_xlsx)
        
        if caminho_pdf:
            conn = get_conn()
            cursor = conn.cursor()
            cursor.execute('''
                UPDATE faturas SET caminho_pdf = ? WHERE id = ?
            ''', (caminho_pdf, fat_id))
            conn.commit()
            conn.close()

if __name__ == "__main__":
    converter_todos_bms()