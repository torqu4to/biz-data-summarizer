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
    print(f"{Cores.VERDE}(2) Recebimentos e tarifas{Cores.RESET}")
    print(f"{Cores.VERDE}(3) Outras transações (não pareadas){Cores.RESET}")
    print(f"{Cores.VERDE}(4) Detalhes por tipo de operação{Cores.RESET}")
    print(f"{Cores.VERDE}(5) Ticket médio diário{Cores.RESET}")
    print(f"{Cores.AMARELO}....................................{Cores.RESET}")
    print(f"{Cores.VERDE}(6) Resumo financeiro{Cores.RESET}")
    print(f"{Cores.VERDE}(7) Gerar relatório completo{Cores.RESET}")
    print(f"{Cores.AMARELO}....................................{Cores.RESET}")
    print(f"{Cores.VERMELHO}(0) Sair{Cores.RESET}")

def gerar_resumo_completo(df, f=None):
    """
    Gera um resumo completo com informações financeiras e estatísticas.
    """
    # Cálculos básicos
    total_recebimentos = df[df['Tipo'] == 'Recebimento']['Valor'].sum()
    total_saidas = abs(df[df['Tipo'].isin(['Saque', 'Transferência', 'Pagamento'])]['Valor'].sum())
    saldo_final = total_recebimentos - total_saidas
    
    # Cálculos de tarifas
    total_tarifas = abs(df[df['Tipo'] == 'Tarifa do Mercado Pago']['Valor'].sum())
    recebimentos_com_relacao = df[(df['Tipo'] == 'Recebimento') & (df['Operacao_Relacionada'].notna())]
    tarifas_com_relacao = df[(df['Tipo'] == 'Tarifa do Mercado Pago') & (df['Operacao_Relacionada'].notna())]
    operacoes_completas = set(recebimentos_com_relacao['Operacao_Relacionada']) & set(tarifas_com_relacao['Operacao_Relacionada'])
    tarifas_pareadas = tarifas_com_relacao[tarifas_com_relacao['Operacao_Relacionada'].isin(operacoes_completas)]
    total_tarifas_pareadas = abs(tarifas_pareadas['Valor'].sum())
    
    # Estatísticas gerais
    periodo_dias = (df['Data'].max() - df['Data'].min()).days + 1
    total_operacoes = len(df)
    media_operacoes_dia = total_operacoes / periodo_dias
    
    # Análise de tarifas
    media_tarifa = total_tarifas_pareadas / len(tarifas_pareadas) if len(tarifas_pareadas) > 0 else 0
    percentual_tarifas = (total_tarifas_pareadas / total_recebimentos * 100) if total_recebimentos > 0 else 0
    
    output = [
        "=== RESUMO FINANCEIRO COMPLETO ===\n",
        f"Período analisado: {periodo_dias} dias",
        f"Total de operações: {total_operacoes}",
        f"Média de operações por dia: {media_operacoes_dia:.1f}\n",
        
        "=== VALORES TOTAIS ===",
        f"Total de Recebimentos: R$ {total_recebimentos:.2f}",
        f"Total de Saídas: R$ {total_saidas:.2f}",
        f"Saldo Final: R$ {saldo_final:.2f}\n",
        
        "=== ANÁLISE DE TARIFAS ===",
        f"Total de Tarifas: R$ {total_tarifas:.2f}",
        f"Total de Tarifas Pareadas: R$ {total_tarifas_pareadas:.2f}",
        f"Média por Tarifa: R$ {media_tarifa:.2f}",
        f"Percentual de Tarifas sobre Recebimentos: {percentual_tarifas:.2f}%"
    ]
    
    if f:
        f.write("\n".join(output))
    else:
        print("\n".join(output))

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
    
    output = [
        "=== ANÁLISE DE RECEBIMENTOS E TARIFAS PAREADOS ===\n",
        f"Quantidade de Recebimentos Pareados: {len(recebimentos_pareados)}",
        f"Quantidade de Tarifas Pareadas: {len(tarifas_pareadas)}",
        f"Total de Recebimentos Pareados: R$ {total_recebimentos_pareados:.2f}",
        f"Total de Tarifas Pareadas: R$ {total_tarifas_pareadas:.2f}",
        f"Saldo (Recebimentos - Tarifas): R$ {saldo_recebimentos_tarifas:.2f}\n"
    ]
    
    if f:
        f.write("\n".join(output))
    else:
        print("\n".join(output))

def gerar_operacoes_nao_pareadas(df, f=None):
    """
    Gera a lista de operações não pareadas.
    """
    recebimentos = df[df['Tipo'] == 'Recebimento']
    tarifas = df[df['Tipo'] == 'Tarifa do Mercado Pago']
    
    recebimentos_com_relacao = recebimentos[recebimentos['Operacao_Relacionada'].notna()]
    tarifas_com_relacao = tarifas[tarifas['Operacao_Relacionada'].notna()]
    
    operacoes_completas = set(recebimentos_com_relacao['Operacao_Relacionada']) & set(tarifas_com_relacao['Operacao_Relacionada'])
    
    recebimentos_nao_pareados = recebimentos_com_relacao[~recebimentos_com_relacao['Operacao_Relacionada'].isin(operacoes_completas)]
    tarifas_nao_pareadas = tarifas_com_relacao[~tarifas_com_relacao['Operacao_Relacionada'].isin(operacoes_completas)]
    
    output = ["=== OPERAÇÕES NÃO PAREADAS ===\n"]
    
    if len(recebimentos_nao_pareados) > 0:
        output.append("\nRecebimentos sem Tarifa Correspondente:")
        for _, row in recebimentos_nao_pareados.iterrows():
            output.append(f"Data: {row['Data'].strftime('%d/%m/%Y')} - Movimento {row['Descrição']} - Operação Relacionada: {row['Operacao_Relacionada']} - Valor: R$ {row['Valor']:.2f}")
    
    if len(tarifas_nao_pareadas) > 0:
        output.append("\nTarifas sem Recebimento Correspondente:")
        for _, row in tarifas_nao_pareadas.iterrows():
            output.append(f"Data: {row['Data'].strftime('%d/%m/%Y')} - Movimento {row['Descrição']} - Operação Relacionada: {row['Operacao_Relacionada']} - Valor: R$ {abs(row['Valor']):.2f}")
    
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

def gerar_ticket_medio_diario(df, f=None):
    """
    Gera o ticket médio por dia da semana baseado em recebimentos e tarifas pareados.
    """
    recebimentos = df[df['Tipo'] == 'Recebimento']
    tarifas = df[df['Tipo'] == 'Tarifa do Mercado Pago']
    
    recebimentos_com_relacao = recebimentos[recebimentos['Operacao_Relacionada'].notna()]
    tarifas_com_relacao = tarifas[tarifas['Operacao_Relacionada'].notna()]
    
    operacoes_completas = set(recebimentos_com_relacao['Operacao_Relacionada']) & set(tarifas_com_relacao['Operacao_Relacionada'])
    
    recebimentos_pareados = recebimentos_com_relacao[recebimentos_com_relacao['Operacao_Relacionada'].isin(operacoes_completas)]
    tarifas_pareadas = tarifas_com_relacao[tarifas_com_relacao['Operacao_Relacionada'].isin(operacoes_completas)]
    
    # Adiciona coluna com o dia da semana
    recebimentos_pareados['Dia_Semana'] = recebimentos_pareados['Data'].dt.day_name()
    tarifas_pareadas['Dia_Semana'] = tarifas_pareadas['Data'].dt.day_name()
    
    # Agrupa por dia da semana
    recebimentos_por_dia = recebimentos_pareados.groupby('Dia_Semana')['Valor'].sum()
    tarifas_por_dia = tarifas_pareadas.groupby('Dia_Semana')['Valor'].sum()
    
    # Combina os dados em um DataFrame
    ticket_medio_diario = pd.DataFrame({
        'Recebimentos': recebimentos_por_dia,
        'Tarifas': abs(tarifas_por_dia)
    }).fillna(0)
    
    # Calcula o valor líquido por dia da semana
    ticket_medio_diario['Valor_Liquido'] = ticket_medio_diario['Recebimentos'] - ticket_medio_diario['Tarifas']
    
    # Ordena os dias da semana
    dias_semana = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    dias_semana_pt = {
        'Monday': 'Segunda-feira',
        'Tuesday': 'Terça-feira',
        'Wednesday': 'Quarta-feira',
        'Thursday': 'Quinta-feira',
        'Friday': 'Sexta-feira',
        'Saturday': 'Sábado',
        'Sunday': 'Domingo'
    }
    
    ticket_medio_diario = ticket_medio_diario.reindex(dias_semana)
    
    output = [
        "=== TICKET MÉDIO POR DIA DA SEMANA ===\n",
        "Valor líquido por dia da semana (Recebimentos - Tarifas):"
    ]
    
    for dia, row in ticket_medio_diario.iterrows():
        if row['Recebimentos'] > 0 or row['Tarifas'] > 0:  # Só mostra dias que tiveram movimentação
            output.append(f"\n{dias_semana_pt[dia]}:")
            output.append(f"  Recebimentos: R$ {row['Recebimentos']:.2f}")
            output.append(f"  Tarifas: R$ {row['Tarifas']:.2f}")
            output.append(f"  Valor Líquido: R$ {row['Valor_Liquido']:.2f}")
    
    if f:
        f.write("\n".join(output))
    else:
        print("\n".join(output))

def processar_relatorio(caminho_arquivo=None):
    """Processa o arquivo Excel e gera o relatório."""
    while True:
        if not caminho_arquivo:
            caminho_arquivo = selecionar_arquivo()
            if not caminho_arquivo:
                continue
        
        print(f"\n{Cores.AZUL}Carregando arquivo...{Cores.RESET}")
        try:
            # Lê o arquivo Excel
            df = pd.read_excel(caminho_arquivo)
            
            # Renomeia as colunas para facilitar o processamento
            df = df.rename(columns={
                'Data de pagamento': 'Data',
                'Tipo de operação': 'Tipo',
                'Número do movimento': 'Descrição',
                'Operação relacionada': 'Operacao_Relacionada',
                'Valor': 'Valor'
            })
            
            # Verifica se as colunas necessárias existem
            colunas_necessarias = ['Data', 'Tipo', 'Valor']
            colunas_faltantes = [col for col in colunas_necessarias if col not in df.columns]
            
            if colunas_faltantes:
                print(f"\n{Cores.VERMELHO}Erro: As seguintes colunas não foram encontradas no arquivo: {', '.join(colunas_faltantes)}{Cores.RESET}")
                print(f"{Cores.AMARELO}Colunas encontradas: {', '.join(df.columns)}{Cores.RESET}")
                caminho_arquivo = None
                continue
            
            # Converte a coluna de data para datetime
            df['Data'] = pd.to_datetime(df['Data'])
            
            # Converte a coluna de valor para numérico
            df['Valor'] = pd.to_numeric(df['Valor'].astype(str).str.replace(',', '.'), errors='coerce')
            
            print(f"{Cores.VERDE}Arquivo carregado com sucesso!{Cores.RESET}")
            
        except Exception as e:
            print(f"\n{Cores.VERMELHO}Erro ao carregar o arquivo: {str(e)}{Cores.RESET}")
            print(f"{Cores.AMARELO}Detalhes do erro:{Cores.RESET}")
            import traceback
            traceback.print_exc()
            caminho_arquivo = None
            continue
        
        while True:
            exibir_menu()
            opcao = input(f"\n{Cores.AZUL}Escolha uma opção: {Cores.RESET}")
            limpar_tela()
            
            if opcao == '0':
                return
            elif opcao == '1':
                caminho_arquivo = selecionar_arquivo()
                break
            elif opcao == '2':
                gerar_analise_recebimentos_tarifas(df)
                input(f"\n{Cores.AZUL}Pressione Enter para continuar...{Cores.RESET}")
            elif opcao == '3':
                gerar_operacoes_nao_pareadas(df)
                input(f"\n{Cores.AZUL}Pressione Enter para continuar...{Cores.RESET}")
            elif opcao == '4':
                gerar_detalhes_por_tipo(df)
                input(f"\n{Cores.AZUL}Pressione Enter para continuar...{Cores.RESET}")
            elif opcao == '5':
                gerar_ticket_medio_diario(df)
                input(f"\n{Cores.AZUL}Pressione Enter para continuar...{Cores.RESET}")
            elif opcao == '6':
                gerar_resumo_completo(df)
                input(f"\n{Cores.AZUL}Pressione Enter para continuar...{Cores.RESET}")
            elif opcao == '7':
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                nome_arquivo = f"relatorio_completo_{timestamp}.txt"
                print(f"\n{Cores.AZUL}Gerando relatório completo...{Cores.RESET}")
                with open(nome_arquivo, 'w', encoding='utf-8') as f:
                    gerar_resumo_completo(df, f)
                    f.write("\n\n")
                    gerar_analise_recebimentos_tarifas(df, f)
                    f.write("\n\n")
                    gerar_operacoes_nao_pareadas(df, f)
                    f.write("\n\n")
                    gerar_detalhes_por_tipo(df, f)
                    f.write("\n\n")
                    gerar_ticket_medio_diario(df, f)
                print(f"\n{Cores.VERDE}Relatório completo gerado com sucesso: {nome_arquivo}{Cores.RESET}")
                input(f"\n{Cores.AZUL}Pressione Enter para continuar...{Cores.RESET}")
            else:
                print(f"\n{Cores.VERMELHO}Opção inválida!{Cores.RESET}")
                input(f"\n{Cores.AZUL}Pressione Enter para continuar...{Cores.RESET}")

if __name__ == "__main__":
    processar_relatorio()