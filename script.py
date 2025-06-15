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
    print(f"{Cores.VERDE}(6) Análise de recebimentos e saídas{Cores.RESET}")
    print(f"{Cores.AMARELO}....................................{Cores.RESET}")
    print(f"{Cores.VERDE}(7) Resumo financeiro{Cores.RESET}")
    print(f"{Cores.VERDE}(8) Gerar relatório completo{Cores.RESET}")
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

def gerar_analise_recebimentos_saidas(df, f=None):
    """
    Verifica recebimentos (já descontando tarifas) e procura por saídas no mesmo dia com o mesmo valor.
    Considera uma margem de tolerância de R$ 0.01 para comparação dos valores.
    """
    # Obtém recebimentos e tarifas
    recebimentos = df[df['Tipo'] == 'Recebimento'].copy()
    tarifas = df[df['Tipo'] == 'Tarifa do Mercado Pago'].copy()
    
    # Obtém saídas (transferências, pagamentos, etc)
    saidas = df[df['Tipo'].isin(['Transferência via Pix', 'Pagamento', 'Saque'])].copy()
    
    # Pareia recebimentos com tarifas
    recebimentos_com_relacao = recebimentos[recebimentos['Operacao_Relacionada'].notna()]
    tarifas_com_relacao = tarifas[tarifas['Operacao_Relacionada'].notna()]
    
    # Cria um DataFrame com recebimentos líquidos
    recebimentos_liquidos = []
    
    # Para cada recebimento, procura sua tarifa correspondente
    for _, recebimento in recebimentos_com_relacao.iterrows():
        tarifa_correspondente = tarifas_com_relacao[
            tarifas_com_relacao['Operacao_Relacionada'] == recebimento['Operacao_Relacionada']
        ]
        
        if not tarifa_correspondente.empty:
            recebimentos_liquidos.append({
                'Data': recebimento['Data'],
                'Valor_Bruto': recebimento['Valor'],
                'Valor_Tarifa': abs(tarifa_correspondente.iloc[0]['Valor']),
                'Valor_Liquido': recebimento['Valor'] - abs(tarifa_correspondente.iloc[0]['Valor']),
                'Descricao_Recebimento': recebimento['Descrição'],
                'Operacao_Relacionada': recebimento['Operacao_Relacionada']
            })
    
    # Converte a lista para DataFrame
    recebimentos_liquidos = pd.DataFrame(recebimentos_liquidos)
    
    # Converte valores para float com 2 casas decimais
    if not recebimentos_liquidos.empty:
        recebimentos_liquidos['Valor_Liquido'] = recebimentos_liquidos['Valor_Liquido'].round(2)
    saidas['Valor'] = saidas['Valor'].round(2)
    
    output = ["=== ANÁLISE DE RECEBIMENTOS E SAÍDAS NO MESMO DIA ===\n"]
    
    # Debug: Mostrar todos os recebimentos líquidos
    output.append("\nDEBUG - Recebimentos Líquidos:")
    for _, rec in recebimentos_liquidos.iterrows():
        output.append(f"Data: {rec['Data'].strftime('%d/%m/%Y %H:%M:%S')} - Valor Bruto: R$ {rec['Valor_Bruto']:.2f} - Tarifa: R$ {rec['Valor_Tarifa']:.2f} - Valor Líquido: R$ {rec['Valor_Liquido']:.2f}")
    
    # Debug: Mostrar todas as saídas
    output.append("\nDEBUG - Saídas:")
    for _, saida in saidas.iterrows():
        output.append(f"Data: {saida['Data'].strftime('%d/%m/%Y %H:%M:%S')} - Valor: R$ {abs(saida['Valor']):.2f}")
    
    output.append("\n=== CORRESPONDÊNCIAS ENCONTRADAS ===")
    
    # Para cada recebimento líquido, procura saídas no mesmo dia
    for _, recebimento in recebimentos_liquidos.iterrows():
        data = recebimento['Data']
        valor_liquido = recebimento['Valor_Liquido']
        
        # Filtra saídas do mesmo dia
        saidas_do_dia = saidas[saidas['Data'].dt.date == data.date()]
        
        if not saidas_do_dia.empty:
            # Procura saídas com valor próximo (considerando margem de tolerância)
            margem_tolerancia = 0.01  # R$ 0.01 de tolerância
            
            # Comparação com arredondamento para 2 casas decimais
            saidas_iguais = saidas_do_dia[
                abs(abs(saidas_do_dia['Valor']) - valor_liquido) <= margem_tolerancia
            ]
            
            if not saidas_iguais.empty:
                output.append(f"\nData: {data.strftime('%d/%m/%Y')}")
                output.append(f"Recebimento:")
                output.append(f"  Movimento: {recebimento['Descricao_Recebimento']}")
                output.append(f"  Operação Relacionada: {recebimento['Operacao_Relacionada']}")
                output.append(f"  Valor Bruto: R$ {recebimento['Valor_Bruto']:.2f}")
                output.append(f"  Tarifa: R$ {recebimento['Valor_Tarifa']:.2f}")
                output.append(f"  Valor Líquido: R$ {valor_liquido:.2f}")
                output.append(f"Saídas correspondentes:")
                
                for _, saida in saidas_iguais.iterrows():
                    output.append(f"  Movimento: {saida['Descrição']}")
                    output.append(f"  Tipo: {saida['Tipo']}")
                    output.append(f"  Valor: R$ {abs(saida['Valor']):.2f}")
                output.append("-" * 50)
    
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
                gerar_operacoes_nao_pareadas(df)
            elif opcao == '4':
                gerar_detalhes_por_tipo(df)
            elif opcao == '5':
                gerar_ticket_medio_diario(df)
            elif opcao == '6':
                gerar_analise_recebimentos_saidas(df)
            elif opcao == '7':
                gerar_resumo_completo(df)
            elif opcao == '8':
                # Gera nome do arquivo baseado na data atual
                data_atual = datetime.now().strftime("%Y%m%d_%H%M%S")
                nome_arquivo = f"relatorio_completo_{data_atual}.txt"
                caminho_completo = os.path.join("reports", nome_arquivo)
                
                with open(caminho_completo, 'w', encoding='utf-8') as f:
                    f.write(f"Relatório gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
                    f.write(f"Arquivo fonte: {os.path.basename(caminho_arquivo)}\n\n")
                    
                    gerar_resumo_completo(df, f)
                    f.write("\n" + "="*50 + "\n\n")
                    gerar_analise_recebimentos_tarifas(df, f)
                    f.write("\n" + "="*50 + "\n\n")
                    gerar_operacoes_nao_pareadas(df, f)
                    f.write("\n" + "="*50 + "\n\n")
                    gerar_detalhes_por_tipo(df, f)
                    f.write("\n" + "="*50 + "\n\n")
                    gerar_ticket_medio_diario(df, f)
                    f.write("\n" + "="*50 + "\n\n")
                    gerar_analise_recebimentos_saidas(df, f)
                
                print(f"\nRelatório completo salvo em: {caminho_completo}")
            else:
                print("\nOpção inválida!")
            
            input("\nPressione Enter para continuar...")
            
    except Exception as e:
        print(f"\nErro ao processar o arquivo: {str(e)}")
        input("\nPressione Enter para continuar...")

if __name__ == "__main__":
    processar_relatorio()