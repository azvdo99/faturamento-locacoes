import sqlite3

DB_PATH = 'data/locacoes.db'

def criar_banco():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS boletins (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            os TEXT NOT NULL,
            numero_bm INTEGER NOT NULL,
            status TEXT,
            data_criacao TEXT,
            data_envio TEXT,
            caminho_xlsx TEXT,
            caminho_pdf TEXT,
            email_message_id TEXT
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS faturas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            numero_fat INTEGER NOT NULL,
            os TEXT NOT NULL,
            bm_id INTEGER NOT NULL,
            status TEXT,
            data_criacao TEXT,
            data_envio TEXT,
            data_vencimento TEXT,
            caminho_xlsx TEXT,
            caminho_pdf TEXT,
            FOREIGN KEY (bm_id) REFERENCES boletins(id)
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS config_os (
            os TEXT PRIMARY KEY,
            titulo_obra TEXT,
            endereco_obra TEXT,
            ativa INTEGER DEFAULT 1
        )
    ''')

    conn.commit()
    conn.close()

    print("Banco criado irmao")

def get_conn():
    return sqlite3.connect(DB_PATH) #facilitar a conexao do banco nas outras functions

if __name__ == "__main__":
        criar_banco()