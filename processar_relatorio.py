import pandas as pd
from datetime import datetime
import os

def processar_relatorio(caminho_arquivo):
    """
    Processa um arquivo Excel contendo relatório financeiro e gera um resumo.
    
    Args:
        caminho_arquivo (str): Caminho para o arquivo Excel
    """
    try:
        # Lê o arquivo Excel
        df = pd.read_excel(caminho_arquivo)
        
        print("\nEstrutura do arquivo:")
        print("Colunas encontradas:", df.columns.tolist())
        print("\nPrimeiras linhas do arquivo:")
        print(df.head())
        
        # Renomeia as colunas para facilitar o processamento
        df = df.rename(columns={
            'Data de pagamento': 'Data',
            'Tipo de operação': 'Tipo',
            'Número do movimento': 'Descrição',
            'Operação relacionada': 'Operacao_Relacionada',
            'Valor': 'Valor'
        })
        
        # Converte a coluna de data para datetime
        df['Data'] = pd.to_datetime(df['Data'])
        
        # Converte a coluna de valor para numérico, removendo possíveis caracteres não numéricos
        df['Valor'] = pd.to_numeric(df['Valor'].astype(str).str.replace(',', '.'), errors='coerce')
        
        # Define os tipos de operação que são entradas
        tipos_entrada = [
            'Rendimento bruto',
            'Recebimento',
            'Adição de dinheiro',
            'Movimentação geral'
        ]
        
        # Define os tipos de operação que são saídas
        tipos_saida = [
            'Imposto de renda',
            'Tarifa do Mercado Pago',
            'Pagamento',
            'Pagamento com desconto recebido',
            'Transferência via Pix'
        ]
        
        # Calcula o resumo financeiro
        entradas = df[df['Tipo'].isin(tipos_entrada)]['Valor'].sum()
        despesas = abs(df[df['Tipo'].isin(tipos_saida)]['Valor'].sum())  # Usa abs() para garantir que as despesas sejam positivas
        saldo = entradas - despesas
        
        # Análise específica de Recebimentos e Tarifas
        recebimentos = df[df['Tipo'] == 'Recebimento']
        tarifas = df[df['Tipo'] == 'Tarifa do Mercado Pago']
        
        # Filtra apenas operações que têm operação relacionada
        recebimentos_com_relacao = recebimentos[recebimentos['Operacao_Relacionada'].notna()]
        tarifas_com_relacao = tarifas[tarifas['Operacao_Relacionada'].notna()]
        
        # Encontra operações relacionadas que têm tanto recebimento quanto tarifa
        operacoes_completas = set(recebimentos_com_relacao['Operacao_Relacionada']) & set(tarifas_com_relacao['Operacao_Relacionada'])
        
        # Filtra operações que têm recebimento e tarifa
        recebimentos_pareados = recebimentos_com_relacao[recebimentos_com_relacao['Operacao_Relacionada'].isin(operacoes_completas)]
        tarifas_pareadas = tarifas_com_relacao[tarifas_com_relacao['Operacao_Relacionada'].isin(operacoes_completas)]
        
        # Encontra operações não pareadas
        recebimentos_nao_pareados = recebimentos_com_relacao[~recebimentos_com_relacao['Operacao_Relacionada'].isin(operacoes_completas)]
        tarifas_nao_pareadas = tarifas_com_relacao[~tarifas_com_relacao['Operacao_Relacionada'].isin(operacoes_completas)]
        
        # Encontra operações sem relação
        recebimentos_sem_relacao = recebimentos[recebimentos['Operacao_Relacionada'].isna()]
        tarifas_sem_relacao = tarifas[tarifas['Operacao_Relacionada'].isna()]
        
        total_recebimentos_pareados = recebimentos_pareados['Valor'].sum()
        total_tarifas_pareadas = abs(tarifas_pareadas['Valor'].sum())
        saldo_recebimentos_tarifas = total_recebimentos_pareados - total_tarifas_pareadas
        
        print("\nAnálise de Recebimentos e Tarifas Pareados:")
        print(f"Quantidade de Recebimentos Pareados: {len(recebimentos_pareados)}")
        print(f"Quantidade de Tarifas Pareadas: {len(tarifas_pareadas)}")
        print(f"Total de Recebimentos Pareados: R$ {total_recebimentos_pareados:.2f}")
        print(f"Total de Tarifas Pareadas: R$ {total_tarifas_pareadas:.2f}")
        print(f"Saldo (Recebimentos - Tarifas): R$ {saldo_recebimentos_tarifas:.2f}")
        
        # Gera o relatório
        data_atual = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        nome_arquivo_saida = f"resumo_financeiro_{data_atual}.txt"
        
        with open(nome_arquivo_saida, 'w', encoding='utf-8') as f:
            f.write("=== RESUMO FINANCEIRO ===\n\n")
            f.write(f"Período analisado: {df['Data'].min().strftime('%d/%m/%Y')} a {df['Data'].max().strftime('%d/%m/%Y')}\n")
            f.write(f"Total de Entradas: R$ {entradas:.2f}\n")
            f.write(f"Total de Despesas: R$ {despesas:.2f}\n")
            f.write(f"Saldo Final: R$ {saldo:.2f}\n\n")
            
            # Adiciona análise de Recebimentos e Tarifas Pareados
            f.write("=== ANÁLISE DE RECEBIMENTOS E TARIFAS PAREADOS ===\n")
            f.write(f"Quantidade de Recebimentos Pareados: {len(recebimentos_pareados)}\n")
            f.write(f"Quantidade de Tarifas Pareadas: {len(tarifas_pareadas)}\n")
            f.write(f"Total de Recebimentos Pareados: R$ {total_recebimentos_pareados:.2f}\n")
            f.write(f"Total de Tarifas Pareadas: R$ {total_tarifas_pareadas:.2f}\n")
            f.write(f"Saldo (Recebimentos - Tarifas): R$ {saldo_recebimentos_tarifas:.2f}\n\n")
            
            # Lista operações sem relação
            if len(recebimentos_sem_relacao) > 0 or len(tarifas_sem_relacao) > 0:
                f.write("=== OPERAÇÕES SEM RELAÇÃO ===\n")
                
                if len(recebimentos_sem_relacao) > 0:
                    f.write("\nRecebimentos sem Operação Relacionada:\n")
                    for _, row in recebimentos_sem_relacao.iterrows():
                        f.write(f"Data: {row['Data'].strftime('%d/%m/%Y')} - Movimento {row['Descrição']} - Valor: R$ {row['Valor']:.2f}\n")
                
                if len(tarifas_sem_relacao) > 0:
                    f.write("\nTarifas sem Operação Relacionada:\n")
                    for _, row in tarifas_sem_relacao.iterrows():
                        f.write(f"Data: {row['Data'].strftime('%d/%m/%Y')} - Movimento {row['Descrição']} - Valor: R$ {abs(row['Valor']):.2f}\n")
            
            # Lista operações não pareadas
            if len(recebimentos_nao_pareados) > 0 or len(tarifas_nao_pareadas) > 0:
                f.write("\n=== OPERAÇÕES NÃO PAREADAS ===\n")
                
                if len(recebimentos_nao_pareados) > 0:
                    f.write("\nRecebimentos sem Tarifa Correspondente:\n")
                    for _, row in recebimentos_nao_pareados.iterrows():
                        f.write(f"Data: {row['Data'].strftime('%d/%m/%Y')} - Movimento {row['Descrição']} - Operação Relacionada: {row['Operacao_Relacionada']} - Valor: R$ {row['Valor']:.2f}\n")
                
                if len(tarifas_nao_pareadas) > 0:
                    f.write("\nTarifas sem Recebimento Correspondente:\n")
                    for _, row in tarifas_nao_pareadas.iterrows():
                        f.write(f"Data: {row['Data'].strftime('%d/%m/%Y')} - Movimento {row['Descrição']} - Operação Relacionada: {row['Operacao_Relacionada']} - Valor: R$ {abs(row['Valor']):.2f}\n")
            
            # Adiciona estatísticas adicionais
            f.write("\n=== ESTATÍSTICAS ADICIONAIS ===\n")
            f.write(f"Número total de transações: {len(df)}\n")
            f.write(f"Número de entradas: {len(df[df['Tipo'].isin(tipos_entrada)])}\n")
            f.write(f"Número de saídas: {len(df[df['Tipo'].isin(tipos_saida)])}\n")
            
            # Adiciona detalhes por tipo de operação
            f.write("\n=== DETALHES POR TIPO DE OPERAÇÃO ===\n")
            for tipo in df['Tipo'].unique():
                total = df[df['Tipo'] == tipo]['Valor'].sum()
                if tipo in tipos_saida:
                    f.write(f"{tipo}: R$ {abs(total):.2f}\n")
                else:
                    f.write(f"{tipo}: R$ {total:.2f}\n")
        
        print(f"\nRelatório gerado com sucesso: {nome_arquivo_saida}")
        return nome_arquivo_saida
        
    except Exception as e:
        print(f"Erro ao processar o arquivo: {str(e)}")
        return None

if __name__ == "__main__":
    # Caminho do arquivo na pasta files
    caminho_arquivo = os.path.join("files", "1581239272_movements_-2025-06-07-195120.xlsx")
    
    if os.path.exists(caminho_arquivo):
        print(f"Processando arquivo: {caminho_arquivo}")
        processar_relatorio(caminho_arquivo)
    else:
        print("Arquivo não encontrado! Verifique se o arquivo está na pasta 'files'.")