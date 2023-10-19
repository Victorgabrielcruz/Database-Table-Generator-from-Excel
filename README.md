# Database-Table-Generator-from-Excel
# Script Python para Criação de Tabela em Banco de Dados a partir de um Arquivo Excel

Este script Python foi desenvolvido para automatizar o processo de criação de uma tabela em um banco de dados a partir dos dados de um arquivo Excel. O programa oferece opções de configuração para escolher o tipo de autenticação com o banco de dados e o arquivo Excel a ser lido.

## Requisitos

- Python 3.x
- Bibliotecas Python:
  - pandas
  - pyodbc

## Uso

1. Execute o script em um ambiente Python.

2. Escolha o tipo de estrutura de banco de dados que deseja usar (autenticação do Windows ou autenticação padrão).

3. Forneça o caminho para o arquivo Excel que contém os dados a serem inseridos na tabela do banco de dados.

4. O script realizará as seguintes etapas:
   - Verificará se a tabela já existe no banco de dados e a apagará, se necessário.
   - Criará a tabela no banco de dados com base nas colunas do arquivo Excel.
   - Inserir os dados do arquivo Excel na tabela do banco de dados.

5. O script registrará eventos, incluindo mensagens de sucesso e erro, no arquivo de log "log_file.txt".

## Funções

### 1. `create_table(connection, table_name, df)`

Esta função cria uma tabela no banco de dados com base nas colunas do DataFrame fornecido. Os tipos de dados das colunas são determinados dinamicamente.

### 2. `insert_table(connection, table_name, df)`

Esta função insere os dados do DataFrame na tabela criada no banco de dados. Os dados são inseridos linha por linha.

### 3. `Verifica_Tabela_Bd(connection, table_name)`

Esta função verifica se a tabela já existe no banco de dados e, se existir, a apaga.

### 4. `ConfigBD()`

Esta função permite configurar a conexão com o banco de dados, usando autenticação do Windows.

### 5. `ConfigBD_Token()`

Esta função permite configurar a conexão com o banco de dados, usando autenticação padrão com nome de usuário e senha.

### 6. `ConfigTable()`

Esta função permite configurar o caminho para o arquivo Excel que será lido para preencher a tabela no banco de dados.

## Autor

Víctor Gabriel Cruz Pereira

## Contato

victorgabrielcruzpereira21@gmail.com

v.g.pplayer21@gmail.com

vgcp@aluno.ifnmg.edu.br
