import pandas as pd
from datetime import datetime
from colorama import init
import shutil
import os
import time

from Componentes.conversor import estados_siglas

from Componentes.Insert import (
    insert_data_to_db_clientes,
    insert_data_to_db_loja,
    insert_data_to_db_produtos,
    insert_data_to_db_fornecedor,
    insert_data_to_db_vendas
)

init()

def parse_int(value):
    try:
        return int(value)
    except (ValueError, TypeError):
        print(f"[ERRO] Valor inválido para inteiro: {value}")
        return None

def parse_float(value):
    try:
        value = str(value).replace("$", "").replace("R$", "").replace(",", ".").strip()
        return float(value)
    except (ValueError, TypeError):
        print(f"[ERRO] Valor inválido para float: {value}")
        return None

def processar_clientes(colunas):
    for i, estado in enumerate(colunas['Estado']):
        if estado in estados_siglas:
            colunas['Estado'][i] = estados_siglas[estado]
    for i in range(len(colunas['nome'])):
        status = insert_data_to_db_clientes(
            colunas['nome'][i],
            colunas['email'][i],
            colunas['senha'][i],
            colunas['cidade'][i],
            colunas['Estado'][i],
            colunas['pais'][i]
        )
        if status != 200:
            print(f"Erro ao inserir linha {i} em clientes.")
            return status
    return 200

def processar_lojas(colunas):
    for i, estado in enumerate(colunas['estado']):
        if estado in estados_siglas:
            colunas['estado'][i] = estados_siglas[estado]
    for i in range(len(colunas['nome_loja'])):
        status = insert_data_to_db_loja(
            colunas['nome_loja'][i],
            colunas['cidade'][i],
            colunas['estado'][i],
            colunas['pais'][i],
            colunas['gerente'][i]
        )
        if status != 200:
            print(f"Erro ao inserir linha {i} em lojas.")
            return status
    return 200

def processar_produtos(colunas):
    for i in range(len(colunas['nome_produto'])):
        fornecedor_id = parse_int(colunas['fornecedor_id'][i])
        preco = parse_float(colunas['Preco'][i])
        if fornecedor_id is None or preco is None:
            print(f"[ERRO] Linha {i}: Dados inválidos (fornecedor_id: {colunas['fornecedor_id'][i]}, preco: {colunas['Preco'][i]})")
            return 400
        status = insert_data_to_db_produtos(
            colunas['nome_produto'][i],
            fornecedor_id,
            preco,
            colunas['descricao'][i]
        )
        if status != 200:
            print(f"Erro ao inserir linha {i} em produtos.")
            return status
    return 200

def processar_fornecedores(colunas):
    for i, estado in enumerate(colunas['Estado']):
        if estado in estados_siglas:
            colunas['Estado'][i] = estados_siglas[estado]
    for i in range(len(colunas['nome_fornecedor'])):
        status = insert_data_to_db_fornecedor(
            colunas['nome_fornecedor'][i],
            colunas['telefone'][i],
            colunas['cidade'][i],
            colunas['Estado'][i],
            colunas['pais'][i]
        )
        if status != 200:
            print(f"Erro ao inserir linha {i} em fornecedores.")
            return status
    return 200

def processar_vendas(colunas):
    for i in range(len(colunas['cliente_id'])):
        quantidade = parse_int(colunas['quantidade'][i])
        valor_total = parse_float(colunas['valor_total'][i])
        if quantidade is None or valor_total is None:
            print(f"[ERRO] Linha {i}: Dados inválidos para quantidade ou valor_total")
            return 400
        try:
            data_formatada = datetime.strptime(colunas['data_venda'][i], '%m/%d/%Y').strftime('%Y-%m-%d')
        except Exception as e:
            print(f"[ERRO] Linha {i}: formato de data inválido - {colunas['data_venda'][i]} - {e}")
            return 400
        status = insert_data_to_db_vendas(
            colunas['cliente_id'][i],
            colunas['produto_id'][i],
            data_formatada,
            quantidade,
            valor_total
        )
        if status != 200:
            print(f"Erro ao inserir linha {i} em vendas.")
            return status
    return 200

def processar_excel():
    Start = time.time()
    print("Iniciando o processamento dos arquivos Excel...")

    arquivos_funcoes = {
        'clientes.xlsx': processar_clientes,
        'lojas.xlsx': processar_lojas,
        'produtos.xlsx': processar_produtos,
        'fornecedor.xlsx': processar_fornecedores,
        'vendas.xlsx': processar_vendas
    }

    for arquivo, funcao in arquivos_funcoes.items():
        origem = f'C:/Users/jamir.rodrigues/Documents/processar/{arquivo}'
        destino = f'C:/Users/jamir.rodrigues/Documents/Processando/{arquivo}'

        if not os.path.exists(origem):
            print(f"\033[1;31mArquivo {arquivo} não encontrado na pasta de origem.\033[m")
            continue

        shutil.move(origem, destino)
        print(f"\033[0;32mArquivo {arquivo} movido para Processando.\033[m")

        df = pd.read_excel(destino, sheet_name='data')
        colunas_em_listas = {col: df[col].tolist() for col in df.columns}

        status = funcao(colunas_em_listas)

        if status == 200:
            destino_final = f'C:/Users/jamir.rodrigues/Documents/Processados/{arquivo}'
            shutil.move(destino, destino_final)
            print(f"\033[7;33;46mArquivo {arquivo} processado com sucesso e movido para Processados.\033[m")
        else:
            destino_erro = f'C:/Users/jamir.rodrigues/Documents/Erros/{arquivo}'
            shutil.move(destino, destino_erro)
            print(f"\033[1;31mArquivo {arquivo} teve erros e foi movido para Erros.\033[m")

    End = time.time()
    print(f"Processamento concluído em {End - Start:.2f} segundos.")

if __name__ == '__main__':
    processar_excel()
    input()
