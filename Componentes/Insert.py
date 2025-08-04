from db import *

def executar_insert(query, valores, entidade):
    conn = get_connection()
    if not conn:
        print(f"[ERRO] Conexão com o banco falhou ao inserir em {entidade}.")
        return 400

    try:
        cursor = conn.cursor()
        cursor.execute(query, valores)
        conn.commit()
        return 200
    except Error as e:
        print(f"[ERRO] Falha ao inserir em {entidade}: {e}")
        return 400
    finally:
        cursor.close()
        conn.close()

def insert_data_to_db_clientes(nome, email, senha, cidade, estado, pais):
    query = """
        INSERT INTO clientes (nome, email, senha, cidade, estado, pais)
        VALUES (?, ?, ?, ?, ?, ?)
    """
    valores = (nome, email, senha, cidade, estado, pais)
    return executar_insert(query, valores, "clientes")

def insert_data_to_db_loja(nome_loja, cidade, estado, pais, gerente):
    query = """
        INSERT INTO lojas (nome_loja, cidade, estado, pais, gerente)
        VALUES (?, ?, ?, ?, ?)
    """
    valores = (nome_loja, cidade, estado, pais, gerente)
    return executar_insert(query, valores, "lojas")

def insert_data_to_db_produtos(nome_produto, fornecedor_id, preco, descricao):
    query = """
        INSERT INTO produtos (nome_produto, fornecedor_id, preco, descricao)
        VALUES (?, ?, ?, ?)
    """
    valores = (nome_produto, fornecedor_id, preco, descricao)
    return executar_insert(query, valores, "produtos")

def insert_data_to_db_fornecedor(nome_fornecedor, telefone, cidade, estado, pais):
    query = """
        INSERT INTO fornecedores (nome_fornecedor, telefone, cidade, estado, pais)
        VALUES (?, ?, ?, ?, ?)
    """
    valores = (nome_fornecedor, telefone, cidade, estado, pais)
    return executar_insert(query, valores, "fornecedores")

def insert_data_to_db_vendas(cliente_id, produto_id, data_venda, quantidade, valor_total):
    query = """
        INSERT INTO vendas (cliente_id, produto_id, data_venda, quantidade, valor_total)
        VALUES (?, ?, ?, ?, ?)
    """
    valores = (cliente_id, produto_id, data_venda, quantidade, valor_total)
    return executar_insert(query, valores, "vendas")
