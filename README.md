# SISTEMA DE FATURAMENTO — LOCADORA EXEMPLO

Esse projeto automatiza o faturamento das locações que eu faço todo mês. Eu trabalho com logística e antes fazia tudo isso na mão, o que tomava muito tempo.

O sistema faz a parte repetitiva sozinho. Eu só preciso entrar para tomar as decisões baseadas nos BMs que já foram aprovados e fazer a gestão dos pedidos (PC) que vão nas faturas. O objetivo é não perder tempo com processo mecânico que o computador resolve. 

É uma ferramenta que eu uso todo mês na prática

## Funcionalidades

- Geração de BMs: Cria os Boletins de Medição a partir da planilha base.
- PDF: Converte as planilhas para PDF.
- E-mail: Envia os e-mails para os contatos cadastrados.
- Aprovações: O sistema lê as respostas de e-mail e atualiza o status dos BMs.
- Pedidos (PC): Gestão e indicação dos pedidos nas faturas.

## Pré-requisitos

- Python 3.x
- Microsoft Excel instalado (necessário para a conversão de PDF)
- Configuração de e-mail (SMTP e IMAP)

## Como Usar

1. Instalar as dependências:
   python -m venv venv
   .\venv\Scripts\activate
   pip install -r requirements.txt

2. Configurar os arquivos na pasta config/:
   Preencher o settings.json, precos.json e o emails_obras.json com os dados da operação.

3. Rodar o sistema:
   python main.py

## Estrutura das Pastas

- main.py: Menu de controle do sistema.
- src/: Código que faz o sistema funcionar (banco, excel, e-mail).
- config/: Onde ficam os arquivos de configuração do usuário.
- data/: Banco de dados das locações.
- templates/: Modelos de Excel usados como base para os documentos.
- faturamento/: Onde os arquivos gerados são salvos mensalmente.
