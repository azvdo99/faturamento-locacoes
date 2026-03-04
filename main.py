import sys
from src.bm import gerar_todos_bms
from src.pdf import converter_todos_bms, converter_todas_faturas
from src.envio_email import enviar_todos_bms, enviar_todas_faturas
from src.fatura import gerar_todas_faturas, buscar_ultimo_fat
from src.email_aprovacao import verificar_respostas, aprovar_bms_manual
from src.ind_orders import main as indicar_pedidos

def input_seguro(mensagem, tipo=str, opcoes_validas=None, permitir_vazio=False):
    while True:
        try:
            valor = input(mensagem).strip()
            
            if not valor and not permitir_vazio:
                print("Campo não pode ser vazio. Tente novamente.")
                continue
            
            if not valor and permitir_vazio:
                return valor
            
            valor = tipo(valor)
            
            if opcoes_validas and valor not in opcoes_validas:
                print(f"Opção inválida. Escolha entre: {opcoes_validas}")
                continue
            
            return valor
            
        except ValueError:
            print(f"Entrada inválida. Digite um valor do tipo {tipo.__name__}.")
        except KeyboardInterrupt:
            print("\n\n  Operação cancelada.")
            return None

def executar_com_protecao(funcao, nome_operacao):
    try:
        funcao()
    except KeyboardInterrupt:
        print(f"\n {nome_operacao} interrompido pelo usuário.")
    except Exception as e:
        print(f"\n  Erro em {nome_operacao}: {e}")
        print("  O menu continuará normalmente.")

def pedir_numero_fatura():
    print("\nÚltimo número de fatura usado (ou Enter para buscar no banco):")
    ultimo = input("  → ").strip()
    
    if not ultimo:
        numero = buscar_ultimo_fat() + 1
        print(f"  Último encontrado no banco. Iniciando em: {numero}")
        return numero
    
    try:
        numero = int(ultimo) + 1
        print(f"  Iniciando em: {numero}")
        return numero
    except ValueError:
        print("   Valor inválido. Buscando no banco...")
        numero = buscar_ultimo_fat() + 1
        print(f"  Iniciando em: {numero}")
        return numero

def menu():
    while True:
        try:
            print("\n" + "=" * 60)
            print("  SISTEMA DE FATURAMENTO — LOCADORA EXEMPLO")
            print("=" * 60)
            
            print("\n  BOLETINS DE MEDIÇÃO:")
            print("   1. Gerar BMs")
            print("   2. Converter BMs para PDF")
            print("   3. Enviar BMs por email")
            print("   4. Verificar aprovações por email")
            print("   5. Aprovar BMs manualmente")
            
            print("\n  FATURAS:")
            print("   6. Gerar Faturas")
            print("   7. Indicar Pedidos")
            print("   8. Converter Faturas para PDF")
            print("   9. Enviar Faturas por email")
            
            print("\n  FLUXOS RÁPIDOS:")
            print("  10. BMs Completo  (1 → 2 → 3)")
            print("  11. Faturas Completo  (6 → 7 → 8 → 9)")
            
            print("\n   0. Sair")
            print("=" * 60)
            
            opcoes_validas = [str(i) for i in range(12)]
            escolha = input_seguro("\n  Escolha uma opção: ", tipo=str, opcoes_validas=opcoes_validas)
            
            if escolha is None:
                continue
            
            if escolha == '0':
                confirmacao = input("\n  Tem certeza que deseja sair? (s/n): ").strip().lower()
                if confirmacao == 's':
                    print("\n  Até logo!\n")
                    sys.exit(0)
                continue
            
            elif escolha == '1':
                executar_com_protecao(gerar_todos_bms, "Geração de BMs")
            
            elif escolha == '2':
                executar_com_protecao(converter_todos_bms, "Conversão de BMs para PDF")
            
            elif escolha == '3':
                executar_com_protecao(enviar_todos_bms, "Envio de BMs")
            
            elif escolha == '4':
                executar_com_protecao(verificar_respostas, "Verificação de aprovações")
            
            elif escolha == '5':
                executar_com_protecao(aprovar_bms_manual, "Aprovação manual de BMs")
            
            elif escolha == '6':
                numero = pedir_numero_fatura()
                executar_com_protecao(lambda: gerar_todas_faturas(numero), "Geração de Faturas")
            
            elif escolha == '7':
                executar_com_protecao(indicar_pedidos, "Indicação de Pedidos")
            
            elif escolha == '8':
                executar_com_protecao(converter_todas_faturas, "Conversão de Faturas para PDF")
            
            elif escolha == '9':
                executar_com_protecao(enviar_todas_faturas, "Envio de Faturas")
            
            elif escolha == '10':
                print("\n  Fluxo BMs Completo\n")
                executar_com_protecao(gerar_todos_bms, "Geração de BMs")
                input("\n  Pressione Enter para converter PDFs...")
                executar_com_protecao(converter_todos_bms, "Conversão de BMs para PDF")
                input("\n  Pressione Enter para enviar emails...")
                executar_com_protecao(enviar_todos_bms, "Envio de BMs")
                print("\n  Fluxo BMs finalizado!")
            
            elif escolha == '11':
                print("\n  Fluxo Faturas Completo\n")
                numero = pedir_numero_fatura()
                executar_com_protecao(lambda: gerar_todas_faturas(numero), "Geração de Faturas")
                input("\n  Pressione Enter para indicar pedidos...")
                executar_com_protecao(indicar_pedidos, "Indicação de Pedidos")
                input("\n  Pressione Enter para converter PDFs...")
                executar_com_protecao(converter_todas_faturas, "Conversão de Faturas para PDF")
                input("\n  Pressione Enter para enviar emails...")
                executar_com_protecao(enviar_todas_faturas, "Envio de Faturas")
                print("\n  Fluxo Faturas finalizado!")
        
        except KeyboardInterrupt:
            print("\n\n  Use a opção 0 para sair.")
            continue

if __name__ == "__main__":
    menu()