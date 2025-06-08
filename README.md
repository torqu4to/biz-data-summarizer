# Processador de Relatórios Financeiros

Este script processa arquivos Excel contendo relatórios financeiros e gera um resumo detalhado das entradas e despesas.

## Requisitos

- Python 3.7 ou superior
- Bibliotecas Python listadas em `requirements.txt`

## Instalação

1. Clone este repositório
2. Instale as dependências:
```bash
pip install -r requirements.txt
```

## Formato do Arquivo Excel

O arquivo Excel deve conter as seguintes colunas:
- Data: Data da transação
- Descrição: Descrição da transação
- Valor: Valor da transação (número)
- Tipo: "Entrada" ou "Despesa"

## Uso

Execute o script:
```bash
python processar_relatorio.py
```

O script irá:
1. Solicitar o caminho do arquivo Excel
2. Processar os dados
3. Gerar um arquivo de texto com o resumo financeiro

## Saída

O script gera um arquivo de texto contendo:
- Período analisado
- Total de entradas
- Total de despesas
- Saldo final
- As 5 maiores despesas
- As 5 maiores entradas 