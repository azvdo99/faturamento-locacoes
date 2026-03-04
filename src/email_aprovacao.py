import imaplib
import email
from email.header import decode_header
from datetime import datetime
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent.parent))
from src.database import get_conn
from src.config import carregar_config

PALAVRAS_APROVACAO = [
    'aprovado', 'aprovada', 'ok', 'pode seguir', 'pode faturar',
    'de acordo', 'confirmo', 'confirmado', 'autorizado', 'autorizo',
    'pode prosseguir', 'tudo certo', 'correto', 'aceito', 'aceita'
]

def verificar_aprovacao_texto(texto):
    texto_lower = texto.lower()
    for palavra in PALAVRAS_APROVACAO:
        if palavra in texto_lower:
            return True
    return False

def buscar_bms_enviados():
    conn = get_conn()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT id, os, numero_bm, email_message_id, caminho_xlsx
        FROM boletins 
        WHERE status = 'enviado' AND email_message_id IS NOT NULL
    ''')
    bms = cursor.fetchall()
    conn.close()
    return [{'id': b[0], 'os': b[1], 'numero_bm': b[2], 'message_id': b[3], 'caminho_xlsx': b[4]} for b in bms]

def aprovar_bm_banco(bm_id):
    conn = get_conn()
    cursor = conn.cursor()
    cursor.execute('UPDATE boletins SET status = ? WHERE id = ?', ('aprovado', bm_id))
    conn.commit()
    conn.close()

def salvar_anexo_aprovacao(bm, msg):
    if not bm['caminho_xlsx']:
        return
    
    pasta_os = Path(bm['caminho_xlsx']).parent
    pasta_aprovacao = pasta_os / 'aprovacao'
    pasta_aprovacao.mkdir(exist_ok=True)
    
    for part in msg.walk():
        if part.get_content_disposition() == 'attachment':
            nome_arquivo = part.get_filename()
            if nome_arquivo:
                caminho = pasta_aprovacao / nome_arquivo
                with open(caminho, 'wb') as f:
                    f.write(part.get_payload(decode=True))
                print(f"  Anexo salvo: {nome_arquivo}")

def verificar_respostas():
    config = carregar_config()
    imap_config = config['imap']
    
    bms_enviados = buscar_bms_enviados()
    
    if not bms_enviados:
        print("Nenhum BM aguardando aprovação por email.")
        return
    
    print(f"\nVerificando respostas para {len(bms_enviados)} BMs...\n")
    
    try:
        mail = imaplib.IMAP4_SSL(imap_config['host'], imap_config['porta'])
        mail.login(imap_config['usuario'], imap_config['senha'])
        mail.select('INBOX')
        
        aprovados = 0
        
        for bm in bms_enviados:
            _, msgs = mail.search(None, f'(HEADER In-Reply-To "{bm["message_id"]}")')
            
            for num in msgs[0].split():
                _, data = mail.fetch(num, '(RFC822)')
                msg = email.message_from_bytes(data[0][1])
                
                corpo = ''
                if msg.is_multipart():
                    for part in msg.walk():
                        if part.get_content_type() == 'text/plain':
                            corpo += part.get_payload(decode=True).decode(errors='ignore')
                else:
                    corpo = msg.get_payload(decode=True).decode(errors='ignore')
                
                if verificar_aprovacao_texto(corpo):
                    aprovar_bm_banco(bm['id'])
                    salvar_anexo_aprovacao(bm, msg)
                    print(f"BM {bm['numero_bm']:02d} - OS {bm['os']} aprovado automaticamente!")
                    aprovados += 1
                    break
        
        mail.close()
        mail.logout()
        
        if aprovados == 0:
            print("Nenhuma aprovação encontrada. BMs pendentes aguardam sua decisão manual.")
        else:
            print(f"\n{aprovados} BMs aprovados automaticamente!")
            
    except Exception as e:
        print(f"Erro ao verificar emails: {e}")

def aprovar_bms_manual():
    bms = buscar_bms_enviados()
    
    if not bms:
        print("\nNenhum BM aguardando aprovação!")
        return
    
    print("\n" + "=" * 50)
    print("BMs AGUARDANDO APROVACAO")
    print("=" * 50 + "\n")
    
    for i, bm in enumerate(bms, 1):
        print(f"{i}. OS {bm['os']} - BM {bm['numero_bm']:02d}")
    
    print("\n" + "=" * 50)
    escolha = input("\nDigite os numeros para APROVAR (ex: 1,3) ou 'todos': ").strip()
    
    if escolha.lower() == 'todos':
        indices = list(range(1, len(bms) + 1))
    else:
        try:
            indices = [int(x.strip()) for x in escolha.split(',')]
        except:
            print("[AVISO] Entrada invalida!")
            return
    
    aprovados = 0
    for i in indices:
        if 1 <= i <= len(bms):
            aprovar_bm_banco(bms[i - 1]['id'])
            print(f"OS {bms[i-1]['os']} - BM {bms[i-1]['numero_bm']:02d} aprovado!")
            aprovados += 1
    
    print(f"\n{aprovados} BMs aprovados!")

if __name__ == "__main__":
    verificar_respostas()