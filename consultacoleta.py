import pyodbc as p ##Módulo de Acesso SQL
import pandas as pd ##Módulo de manipulação Excel openpyxl
from datetime import datetime, timedelta ##Importando class datetime
import configparser as config ##Módulo de manipulação de arquivos

##Conexão arquivi ini usando a função ConfigParser()
cfg = config.ConfigParser()
cfg.read('config.ini') ##lendo arquivo ini

##Criando variaveis de cada parametro do arquivo ini
instancia = cfg.get("ConexaoSql","instancia") ##(<"topicoDoArquivo">,<"ParametroDoTopico">)
porta = cfg.get("ConexaoSql","porta")
bancoDeDados = cfg.get("ConexaoSql","bancoDeDados")
usuario = cfg.get("ConexaoSql","usuario")
senha = cfg.get("ConexaoSql","senha")

##Conexão ao Banco de Dados, usando os parametros do arquivo ini
cnxn = p.connect("DRIVER={ODBC Driver 17 for SQL Server};"
                      # IP ou nome do servidor.
                      r"SERVER="+instancia+";"
                      # Porta
                      "PORT="+porta+";"
                      # Banco que será utilizado.
                      "DATABASE="+bancoDeDados+";"
                      # Nome de usuário convertido em string.
                      f"UID={usuario};"
                      # Senha convertida em string.
                      f"PWD={senha}"
                      )

##tratando o datetime para formato apenas data e obtendo data atual menos 1 dia
data1 = (datetime.now() - timedelta(days=1))
data1format = data1.strftime("%d%m%Y")

##tratando o datetime para formato apenas data e obtendo a data atual
data2 = datetime.now()
data2format = data2.strftime("%d%m%Y")

localArquivo =  cfg.get("Arquivo_LocalParaGravar","armazenamento")

select = cfg.get("ConsultaSQL","comando")

pd.read_sql(
        'SELECT * FROM {}'.format(select),cnxn
    ).to_excel(
            '{}{}_{}.xlsx'.format(localArquivo,data1format,data2format),'Resultado', index=False
        )

continuar = input("Aperte Enter para Continuar")