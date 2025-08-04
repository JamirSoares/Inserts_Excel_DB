
<body>

<h1>Projeto de Processamento e Inserção de Dados em Banco SQL</h1>

<h2>Descrição</h2>
<p>Este projeto automatiza o processamento de arquivos Excel contendo dados de clientes, lojas, produtos, fornecedores e vendas, e insere esses dados em um banco de dados SQL Server. O sistema organiza o fluxo de arquivos entre pastas de processamento e trata erros de importação, garantindo integridade e padronização, como a abreviação dos nomes dos estados brasileiros.</p>

<h2>Estrutura do Projeto</h2>
<ul>
  <li><strong>db.py</strong><br />Função para conectar ao banco de dados SQL Server.</li>
  <li><strong>Insert.py</strong><br />Funções para inserir dados nas tabelas do banco (clientes, lojas, produtos, fornecedores, vendas) usando uma função comum para executar inserts.</li>
  <li><strong>ProcessarExcel.py</strong><br />Script principal que lê os arquivos Excel, processa seus dados e chama as funções de inserção. Move os arquivos entre pastas de processamento, sucesso ou erro.</li>
  <li><strong>conversor.py</strong><br />Dicionário com a conversão dos nomes completos dos estados brasileiros para siglas.</li>
</ul>

<h2>Como usar</h2>
<ol>
  <li>Configure a conexão no arquivo <code>db.py</code>.</li>
  <li>Coloque os arquivos Excel na pasta:<br />
    <code>C:/Users/jamir.rodrigues/Documents/processar/</code><br />
    com os nomes:
    <ul>
      <li>clientes.xlsx</li>
      <li>lojas.xlsx</li>
      <li>produtos.xlsx</li>
      <li>fornecedor.xlsx</li>
      <li>vendas.xlsx</li>
    </ul>
  </li>
  <li>Execute o script <code>ProcessarExcel.py</code>:<br />
    <pre>python ProcessarExcel.py</pre>
  </li>
  <li>Os arquivos serão movidos automaticamente para as pastas <code>Processando</code>, <code>Processados</code> ou <code>Erros</code>, conforme o resultado do processamento.</li>
</ol>

<h2>Detalhes Técnicos</h2>
<ul>
  <li>Usa pandas para ler arquivos Excel.</li>
  <li>Usa colorama para saída colorida no terminal.</li>
  <li>Converte valores monetários e datas para formatos adequados.</li>
  <li>Usa dicionário <code>estados_siglas</code> para padronizar os estados.</li>
  <li>Modularizado para facilitar manutenção.</li>
  <li>Trata erros de inserção e imprime mensagens claras.</li>
</ul>

<h2>Dependências</h2>
<ul>
  <li>Python 3.x</li>
  <li>pandas</li>
  <li>colorama</li>
  <li>pyodbc (ou driver ODBC para SQL Server)</li>
  <li>openpyxl</li>
</ul>
<p>Instale as dependências via pip:<br />
<pre>pip install pandas colorama pyodbc openpyxl</pre></p>

<h2>Funções Principais</h2>

<h3>Insert.py</h3>
<ul>
  <li><code>executar_insert(query, valores, entidade)</code><br />
  Executa o comando SQL com tratamento de erros.</li>
  <li>Funções de inserção específicas para: clientes, lojas, produtos, fornecedores e vendas.</li>
</ul>

<h3>ProcessarExcel.py</h3>
<ul>
  <li>Funções para processar dados de cada tipo de arquivo: clientes, lojas, produtos, fornecedores e vendas.</li>
  <li>Função principal <code>processar_excel()</code> controla o fluxo completo do processamento.</li>
</ul>

<h3>conversor.py</h3>
<p>Dicionário para conversão dos nomes dos estados brasileiros em suas siglas.</p>

<h2>Observações</h2>
<ul>
  <li>Os arquivos Excel devem conter uma aba chamada <code>data</code>.</li>
  <li>Colunas devem estar com nomes exatos esperados.</li>
  <li>Datas em <code>vendas.xlsx</code> devem estar no formato <code>MM/DD/YYYY</code>.</li>
  <li>Certifique-se de que o banco esteja configurado corretamente para receber os dados.</li>
</ul>

<hr />

<p><em>Desenvolvido por Jamir - System Analytics</em></p>

</body>
</html>
