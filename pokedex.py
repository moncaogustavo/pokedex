import os
from datetime import datetime
import pandas as pd
import oracledb
os.system("cls")

def conectar():
    try:
        conn = oracledb.connect(user="RM557515", password="250803", dsn="oracle.fiap.com.br:1521/ORCL")
        print("Conexão bem-sucedida!")
        return conn
    except oracledb.Error as e:
        print("Erro ao conectar ao bando de dados:", e)
        return None
    
def cadastrar(conn):
    try:
        cursor = conn.cursor()
        nome = input("Digite o nome do Pokémon: ")
        tipo_primario = input("Digite o tipo primário: ")
        tipo_secundario = input("Digite o tipo secundário (deixe em branco se não tiver): ")

        if tipo_secundario == '':
            tipo_secundario = None

        sql = "INSERT INTO pokedex (nome, tipo_primario, tipo_secundario) VALUES (:1, :2, :3)"
        cursor.execute(sql, [nome, tipo_primario, tipo_secundario])
        conn.commit()
        print(f"Pokémon {nome} cadastrado com sucesso!")
    except oracledb.Error as e:
        print("Erro ao cadastrar Pokémon:", e)

def listar_exportar(conn):
    try:
        campos_disponiveis = ['id', 'nome', 'tipo_primario', 'tipo_secundario']
        print("\nCampos disponíveis: id, nome, tipo_primario, tipo_secundario")
        campos_selecionados = input("Digite os campos que deseja visualizar/exportar (separados por vírgula): ").split(',')

        campos_selecionados = [campo.strip() for campo in campos_selecionados if campo.strip() in campos_disponiveis]

        if not campos_selecionados:
            print("Nenhum campo válido selecionado. Tente novamente.")
            return

        sql = f"SELECT {', '.join(campos_selecionados)} FROM pokedex"
        cursor = conn.cursor()
        cursor.execute(sql)
        rows = cursor.fetchall()

        print(f"\nListando Pokémon na Pokédex com os campos: {', '.join(campos_selecionados)}")
        for row in rows:
            print(row)

        exportar = input("\nDeseja exportar os dados para um arquivo Excel (.xlsx)? (s/n): ").lower()
        if exportar == 's':
            df = pd.DataFrame(rows, columns=campos_selecionados)

            nome_arquivo = datetime.now().strftime("planilha%Y%m%d%H%M%S.xlsx")

            df.to_excel(nome_arquivo, index=False, engine='openpyxl')
            print(f"Dados exportados para '{nome_arquivo}' com sucesso!")
    except oracledb.Error as e:
        print("Erro ao listar/exportar Pokémon:", e)

def menu():
    conn = conectar()
    if conn is None:
        return
    
    while True:
        print("\nMENU POKÉDEX:")
        print("0 - SAIR")
        print("1 - CADASTRAR POKÉMON")
        print("2 - LISTAR/EXPORTAR POKÉMON")
        opcao = input("Escolha uma opção: ")

        if opcao == '0':
            print("Saindo da instalação...")
            break
        elif opcao == '1':
            cadastrar(conn)
        elif opcao == '2':
            listar_exportar(conn)
        else:
            print("Opção inválida, tente novamente.")
    if conn:
        conn.close()
menu()
