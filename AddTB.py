import pandas as pd
import pyodbc
import os
import logging
logging.basicConfig(filename='log_file.txt', level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')

# [1.0] Funções

# [1.1] Função que cria a tabela no banco de dados

def create_table(connection, table_name, df):
    # Cria um cursor para executar comandos SQL no banco de dados.
    cursor = connection.cursor()

    # Cria uma string que representa o comando SQL para criar uma tabela com um nome especificado.
    # O nome da tabela é obtido a partir da variável table_name.
    create_table_sql = f"CREATE TABLE {table_name} ("

    # Inicia um loop que percorrerá as colunas no DataFrame df.
    for column in df.columns:

        # Obtém o tipo de dados da coluna atual.
        data_type = df[column].dtype

        # Verifica se o tipo de dados é uma string (object).
        if 'object' in str(data_type):

            # Adiciona à string create_table_sql a definição da coluna como VARCHAR(MAX),
            # seguido por uma vírgula.
            create_table_sql += f"{column} VARCHAR(MAX),"

        # Se o tipo de dados não é uma string, verifica se é um inteiro (int).
        elif 'int' in str(data_type):
            create_table_sql += f"{column} INT, "

        # Se o tipo de dados não é uma string nem um inteiro, verifica se é um ponto flutuante (float).
        elif 'float' in str(data_type):
            create_table_sql += f"{column} FLOAT, "

    # Remove os dois últimos caracteres da string create_table_sql (vírgula e espaço) e adiciona
    # um parêntese de fechamento para concluir a definição da tabela.
    create_table_sql = create_table_sql[:-1] + ")"

    # Executa o comando SQL create_table_sql no banco de dados usando o cursor.
    cursor.execute(create_table_sql)

    # Confirma a transação, garantindo que as alterações sejam permanentes no banco de dados.
    connection.commit()


# -----------------------------------------------------------
# [1.2] Função que inseri dados na tabela
def insert_table(connection, table_name, df):
    try:
        # Cria um cursor para executar comandos SQL no banco de dados.
        cursor = connection.cursor()

        # Inicia um loop que percorrerá cada linha do DataFrame df.
        for _, row in df.iterrows():
            # Cria um placeholder de interrogação para cada valor na linha.
            placeholder = ",".join("?" * len(row))

            # Cria a string do comando SQL de inserção.
            # O nome da tabela é obtido a partir da variável table_name.
            insert_sql = f"INSERT INTO {table_name} VALUES ({placeholder})"

            # Executa o comando SQL de inserção, passando uma tupla com os valores da linha como argumento.
            cursor.execute(insert_sql, tuple(row))

            # Confirma a transação, garantindo que as inserções sejam permanentes no banco de dados.
        connection.commit()

    except Exception as e:
        print("Erro ao inserir os dados na tabela:", str(e))


# ---------------------------------------------------------------------------------------------------------
# [1.3] Função que verfica se a tabela ja existe e a apaga caso ja exista

def Verifica_Tabela_Bd(connection, table_name):
    try:
        # Cria um cursor para executar comandos SQL no banco de dados.
        cursor = connection.cursor()

        # Comando SQL para verificar se a tabela existe na base de dados.
        sql_command = f"SELECT COUNT(*) FROM information_schema.tables WHERE table_name = '{table_name}'"

        # Executa o comando SQL.
        cursor.execute(sql_command)

        # Recupera o resultado da consulta.
        op = cursor.fetchone()[0]

        # Verifica se a tabela existe (se a contagem retornada é igual a 1).
        if op == 1:
            # Se a tabela existe, cria um comando SQL para apagá-la.
            sql_command = f'DROP TABLE {table_name}'
            cursor.execute(sql_command)

        # Confirma as alterações no banco de dados.
        connection.commit()
    except Exception as e:
        print("Erro na verificação:", str(e))
    finally:
        return
# -------------------------------------------------------------------------------------------------
# [1.4] Função que configura o Banco de Dados

# [1.4.1] Configuração do BD com autenticação do Windows
def ConfigBD():
    connection_string = ""
    while True:
        try:
            server_name = input(" Adicione o nome do servidor: ")
            database_name = input(" Adicione o nome do Banco de Dados: ")
            connection_string = f"DRIVER={{SQL Server}};SERVER={server_name};DATABASE={database_name};Trusted_Connection=yes;"
            break
        except Exception as e:
            print(" Erro na conexão com o Banco de dados, por favor tente novamente")
    return connection_string

# -----------------------------------------------------------
# [1.4.2] Configuração do padrão do BD
def ConfigBD_Token():
    connection_string = ""
    while True:
        try:
            server_name = input(" Adicione o nome do servidor: ")
            database_name = input(" Adicione o nome do Banco de Dados: ")
            User = input(" Adicione o User: ")
            password = input(" Adicione a senha: ")
            connection_string = f'DRIVER={{SQL Server}};SERVER={server_name};DATABASE={database_name};UID={User};PWD={password}'
            break
        except Exception as e:
            print("Erro na conexão com o Banco de dados, por favor tente novamente \n", e)
    return connection_string
# -----------------------------------------------------------
# [1.5] Confiuguração de endereço do arquivo excel
def ConfigTable ():

    while True:
        file = None
        try:
            op = int(input("--------------- \n O arquivo excel esta na mesma pasta deste arquivo python ? \n[1] Sim \n[2] Não \n"))
            if op == 1:
                script_directory = os.path.dirname(os.path.abspath(__file__))
                file_name = input("Ensira o nome do arquivo excel: ")
                file_name = f'{file_name}'
                file = os.path.join(script_directory, file_name)
                break
            elif op == 2 :
                script_directory = input(" Inira o caminho do arquivo, como no exemplo C:\ User\SeuNome\Documentos\MeuProjeto\ relatorio.xlsx: \n:")
                script_directory = r'{script_directory}'
                file = script_directory
                break
            else:
                print (" Opção invalida")
        except Exception as e:
            print(" Erro! Tente novamente \n", e)
    return file
# -----------------------------------------------------------
# [2.0] Main

# [2.1] Definindo a conexão ao Banco de Dados

try:
    while True:
        op = int(input("--------------- \n Defina a estrutúra de banco de dados que ira utilizar: \n [1] Banco de Dados com autenticação do windoes \n [2] Banco de Dados padrão \n" ))
        if op == 1:
            connection = pyodbc.connect(ConfigBD())
            break
        elif op == 2 :
            connection = pyodbc.connect(ConfigBD_Token())
            break
        else :
            print  ("[ERRO] - Escolha invalida, tente novamente")
# [2.2] Procurando o arquivo excel
    excel_file = ConfigTable()
    table_name = os.path.splitext(os.path.basename(excel_file))[0]

# [2.3] Lendo a planilha Excel em um DataFrame
    df = pd.read_excel(excel_file)
# [2.4] Verificando o Banco de Dados
    Verifica_Tabela_Bd(connection, table_name)
    logging.info(f'Tabela {table_name} verificada e, se existente, apagada com sucesso.')
# [2.5] Criar a tabela no BD
    create_table(connection, table_name, df)
    logging.info(f'Tabela {table_name} criada com sucesso no banco de dados.')
# [2.6] Insere os dados na tabela
    insert_table(connection, table_name, df)
# [2.7] Fecha a conexão
    connection.close()
    logging.info('Conexão ao banco de dados encerrada com sucesso.')

except :
    print("[ERRO] - Tente novamente, se o erro percistir envie um email para o desenvolvedor \n victorgabrielcruzpereira21@gmail.com")
logging.info('[SUCESS] Código executado com sucesso')

print("Sucesso")
