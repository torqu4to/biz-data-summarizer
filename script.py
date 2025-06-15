import pandas as pd
from datetime import datetime
import os
import glob
import warnings

# Suprime os warnings do openpyxl
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# Códigos de cores ANSI
class Cores:
    VERDE = '\033[92m'
    AZUL = '\033[94m'
    AMARELO = '\033[93m'
    VERMELHO = '\033[91m'
    MAGENTA = '\033[95m'
    RESET = '\033[0m'

def limpar_tela():
    """Limpa a tela do terminal."""
    os.system('cls' if os.name == 'nt' else 'clear')

def listar_arquivos_excel():
    """
    Lista todos os arquivos Excel na pasta 'files'.
    """
    arquivos = glob.glob(os.path.join("files", "*.xlsx"))
    if not arquivos:
        return []
    return arquivos

def validar_arquivo_excel(caminho_arquivo):
    """
    Valida se o arquivo Excel tem a estrutura correta.
    
    Args:
        caminho_arquivo (str): Caminho para o arquivo Excel
        
    Returns:
        tuple: (bool, str) - (é_válido, mensagem_erro)
    """
    try:
        # Tenta ler o arquivo
        df = pd.read_excel(caminho_arquivo)
        
        # Verifica as colunas necessárias
        colunas_necessarias = [
            'Data de pagamento',
            'Tipo de operação',
            'Número do movimento',
            'Operação relacionada',
            'Valor'
        ]
        
        colunas_faltantes = [col for col in colunas_necessarias if col not in df.columns]
        if colunas_faltantes:
            return False, f"Colunas necessárias não encontradas: {', '.join(colunas_faltantes)}"
        
        return True, "Arquivo válido"
        
    except Exception as e:
        return False, f"Erro ao validar arquivo: {str(e)}"

def selecionar_arquivo():
    """
    Permite ao usuário selecionar um arquivo Excel para processar.
    
    Returns:
        str: Caminho do arquivo selecionado ou None se cancelado
    """
    arquivos = listar_arquivos_excel()
    
    if not arquivos:
        print("\nNenhum arquivo Excel encontrado na pasta 'files'!")
        return None
    
    print("\n=== ARQUIVOS DISPONÍVEIS ===")
    for i, arquivo in enumerate(arquivos, 1):
        nome_arquivo = os.path.basename(arquivo)
        print(f"{i}. {nome_arquivo}")
    print("0. Voltar")
    
    while True:
        try:
            opcao = input("\nEscolha um arquivo (número) ou 0 para voltar: ")
            
            if opcao == '0':
                return None
            
            indice = int(opcao) - 1
            if 0 <= indice < len(arquivos):
                arquivo_selecionado = arquivos[indice]
                valido, mensagem = validar_arquivo_excel(arquivo_selecionado)
                
                if valido:
                    return arquivo_selecionado
                else:
                    print(f"\nErro: {mensagem}")
                    return None
            else:
                print("\nOpção inválida! Por favor, escolha um número válido.")
        except ValueError:
            print("\nPor favor, digite um número válido.")

def exibir_menu():
    """Exibe o menu principal."""
    print(f"\n{Cores.AZUL}=== MENU PRINCIPAL ==={Cores.RESET}")
    print(f"{Cores.VERDE}(1) Escolher outro arquivo{Cores.RESET}")
    print(f"{Cores.AMARELO}....................................{Cores.RESET}")
    print(f"{Cores.VERDE}(2) Recebimentos e tarifas (Resumo financeiro){Cores.RESET}")
    print(f"{Cores.VERDE}(3) Detalhes por tipo de operação{Cores.RESET}")
    print(f"{Cores.VERDE}(4) Entradas maiores que R$ 59,00{Cores.RESET}")
    print(f"{Cores.AMARELO}....................................{Cores.RESET}")
    print(f"{Cores.VERDE}(5) Gerar relatório completo{Cores.RESET}")
    print(f"{Cores.AMARELO}....................................{Cores.RESET}")
    print(f"{Cores.VERMELHO}(0) Sair{Cores.RESET}")

def gerar_analise_recebimentos_tarifas(df, f=None):
    """
    Gera a análise de recebimentos e tarifas.
    """
    recebimentos = df[df['Tipo'] == 'Recebimento']
    tarifas = df[df['Tipo'] == 'Tarifa do Mercado Pago']
    
    recebimentos_com_relacao = recebimentos[recebimentos['Operacao_Relacionada'].notna()]
    tarifas_com_relacao = tarifas[tarifas['Operacao_Relacionada'].notna()]
    
    operacoes_completas = set(recebimentos_com_relacao['Operacao_Relacionada']) & set(tarifas_com_relacao['Operacao_Relacionada'])
    
    recebimentos_pareados = recebimentos_com_relacao[recebimentos_com_relacao['Operacao_Relacionada'].isin(operacoes_completas)]
    tarifas_pareadas = tarifas_com_relacao[tarifas_com_relacao['Operacao_Relacionada'].isin(operacoes_completas)]
    
    total_recebimentos_pareados = recebimentos_pareados['Valor'].sum()
    total_tarifas_pareadas = abs(tarifas_pareadas['Valor'].sum())
    saldo_recebimentos_tarifas = total_recebimentos_pareados - total_tarifas_pareadas
    
    # Estatísticas gerais
    periodo_dias = (df['Data'].max() - df['Data'].min()).days + 1
    total_operacoes = len(df)
    media_operacoes_dia = total_operacoes / periodo_dias
    
    # Análise de tarifas
    media_tarifa = total_tarifas_pareadas / len(tarifas_pareadas) if len(tarifas_pareadas) > 0 else 0
    percentual_tarifas = (total_tarifas_pareadas / total_recebimentos_pareados * 100) if total_recebimentos_pareados > 0 else 0
    
    output = [
        "=== RESUMO FINANCEIRO ===\n",
        f"Período analisado: {periodo_dias} dias",
        f"Total de operações: {total_operacoes}",
        f"Média de operações por dia: {media_operacoes_dia:.1f}\n",
        
        "=== RECEBIMENTOS E TARIFAS PAREADOS ===",
        f"Quantidade de Recebimentos Pareados: {len(recebimentos_pareados)}",
        f"Quantidade de Tarifas Pareadas: {len(tarifas_pareadas)}",
        f"Total de Recebimentos Pareados: R$ {total_recebimentos_pareados:.2f}",
        f"Total de Tarifas Pareadas: R$ {total_tarifas_pareadas:.2f}",
        f"Saldo (Recebimentos - Tarifas): R$ {saldo_recebimentos_tarifas:.2f}\n",
        
        "=== ANÁLISE DE TARIFAS ===",
        f"Média por Tarifa: R$ {media_tarifa:.2f}",
        f"Percentual de Tarifas sobre Recebimentos: {percentual_tarifas:.2f}%"
    ]
    
    if f:
        f.write("\n".join(output))
    else:
        print("\n".join(output))

def gerar_detalhes_por_tipo(df, f=None):
    """
    Gera detalhes por tipo de operação.
    """
    tipos_saida = [
        'Imposto de renda',
        'Tarifa do Mercado Pago',
        'Pagamento',
        'Pagamento com desconto recebido',
        'Transferência via Pix'
    ]
    
    output = ["=== DETALHES POR TIPO DE OPERAÇÃO ===\n"]
    
    for tipo in df['Tipo'].unique():
        total = df[df['Tipo'] == tipo]['Valor'].sum()
        if tipo in tipos_saida:
            output.append(f"{tipo}: -R$ {abs(total):.2f}")
        else:
            output.append(f"{tipo}: R$ {total:.2f}")
    
    if f:
        f.write("\n".join(output))
    else:
        print("\n".join(output))

def analisar_entradas_maiores(df, f=None):
    """
    Analisa entradas com valor maior que R$ 59,00, excluindo tarifas.
    Mostra detalhes como data, valor, tarifa relacionada (se houver) e outras informações relevantes.
    """
    # Filtra entradas maiores que R$ 59,00 e que não são tarifas
    entradas_maiores = df[
        (df['Valor'] > 59) & 
        (df['Tipo'] != 'Tarifa do Mercado Pago')
    ].copy()
    
    # Ordena por valor (do maior para o menor)
    entradas_maiores = entradas_maiores.sort_values('Valor', ascending=False)
    
    output = [
        "=== ANÁLISE DE ENTRADAS MAIORES QUE R$ 59,00 ===\n",
        f"Total de entradas encontradas: {len(entradas_maiores)}\n"
    ]
    
    # Para cada entrada, procura a tarifa relacionada (se houver)
    for _, entrada in entradas_maiores.iterrows():
        output.append(f"\nData: {entrada['Data'].strftime('%d/%m/%Y %H:%M:%S')}")
        output.append(f"Tipo: {entrada['Tipo']}")
        output.append(f"Movimento: {entrada['Descrição']}")
        output.append(f"Valor: R$ {entrada['Valor']:.2f}")
        
        # Procura tarifa relacionada
        if pd.notna(entrada['Operacao_Relacionada']):
            tarifa_relacionada = df[
                (df['Tipo'] == 'Tarifa do Mercado Pago') & 
                (df['Operacao_Relacionada'] == entrada['Operacao_Relacionada'])
            ]
            
            if not tarifa_relacionada.empty:
                tarifa = tarifa_relacionada.iloc[0]
                output.append(f"Tarifa Relacionada: R$ {abs(tarifa['Valor']):.2f}")
                output.append(f"Valor Líquido: R$ {(entrada['Valor'] - abs(tarifa['Valor'])):.2f}")
        
        output.append("-" * 50)
    
    # Adiciona estatísticas gerais
    total_entradas = entradas_maiores['Valor'].sum()
    output.append(f"\nTotal de Entradas: R$ {total_entradas:.2f}")
    
    # Calcula média dos valores
    media_entradas = total_entradas / len(entradas_maiores) if len(entradas_maiores) > 0 else 0
    output.append(f"Média dos Valores: R$ {media_entradas:.2f}")
    
    if f:
        f.write("\n".join(output))
    else:
        print("\n".join(output))

def processar_relatorio(caminho_arquivo=None):
    """
    Processa o relatório financeiro.
    """
    if caminho_arquivo is None:
        caminho_arquivo = selecionar_arquivo()
        if caminho_arquivo is None:
            return

    try:
        # Lê o arquivo Excel
        df = pd.read_excel(caminho_arquivo)
        
        # Renomeia as colunas para facilitar o acesso
        df = df.rename(columns={
            'Data de pagamento': 'Data',
            'Tipo de operação': 'Tipo',
            'Número do movimento': 'Descrição',
            'Operação relacionada': 'Operacao_Relacionada',
            'Valor': 'Valor'
        })
        
        # Converte a coluna de data para datetime
        df['Data'] = pd.to_datetime(df['Data'])
        
        while True:
            limpar_tela()
            exibir_menu()
            
            opcao = input("\nEscolha uma opção: ")
            
            if opcao == '0':
                break
            elif opcao == '1':
                novo_arquivo = selecionar_arquivo()
                if novo_arquivo:
                    caminho_arquivo = novo_arquivo
                    df = pd.read_excel(caminho_arquivo)
                    df = df.rename(columns={
                        'Data de pagamento': 'Data',
                        'Tipo de operação': 'Tipo',
                        'Número do movimento': 'Descrição',
                        'Operação relacionada': 'Operacao_Relacionada',
                        'Valor': 'Valor'
                    })
                    df['Data'] = pd.to_datetime(df['Data'])
            elif opcao == '2':
                gerar_analise_recebimentos_tarifas(df)
            elif opcao == '3':
                gerar_detalhes_por_tipo(df)
            elif opcao == '4':
                analisar_entradas_maiores(df)
            elif opcao == '5':
                # Gera nome do arquivo baseado na data atual
                data_atual = datetime.now().strftime("%Y%m%d_%H%M%S")
                nome_arquivo = f"relatorio_completo_{data_atual}.txt"
                caminho_completo = os.path.join("reports", nome_arquivo)
                
                with open(caminho_completo, 'w', encoding='utf-8') as f:
                    f.write(f"Relatório gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
                    f.write(f"Arquivo fonte: {os.path.basename(caminho_arquivo)}\n\n")
                    
                    gerar_analise_recebimentos_tarifas(df, f)
                    f.write("\n" + "="*50 + "\n\n")
                    gerar_detalhes_por_tipo(df, f)
                    f.write("\n" + "="*50 + "\n\n")
                    analisar_entradas_maiores(df, f)
                
                print(f"\nRelatório completo salvo em: {caminho_completo}")
            else:
                print("\nOpção inválida!")
            
            input("\nPressione Enter para continuar...")
            
    except Exception as e:
        print(f"\nErro ao processar o arquivo: {str(e)}")
        input("\nPressione Enter para continuar...")

if __name__ == "__main__":
    processar_relatorio()