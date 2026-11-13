# Instalação e importação das Bibliotecas
#pip install openpyxl
#pip install pandas
#import openpyxl
#import pandas as pd



from datetime import date, timedelta
from openpyxl.styles import Alignment
from datetime import date
from openpyxl import load_workbook
from copy import copy
import pandas as pd
import os

print("PASTA DE TRABALHO:", os.getcwd())
print("ARQUIVO DE EMPRÉSTIMO:", os.path.abspath('Tabelas/Controle de Emprestimo.xlsx'))



#importação dos arquivos (Pesquisa e aulas assistida como montar drive)
controleemprestimo = load_workbook('Tabelas/Controle de Emprestimo.xlsx')
cadastroleitores = load_workbook('Tabelas/Cadastro de Leitores.xlsx')
cadastrolivros = load_workbook('Tabelas/Cadastro de Livros.xlsx')
opcao = 0

# BLOCO 1
# Iniciando o sistema com apresentação do Menu Princial
def menu_principal():
# Mensagem iniciala
  print()
  print()
  print("=" *74)
  print(" Seja Bem-vindos ao Sistema de Gestão da Biblioteca Municipal de Camocim!")
  print("=" *74)

#Apresentação do Menu Inicail
  print()
  print()
  print ("Por favor selecione uma das opções a baixo:")
  
  print()
  print("1 - Cadastro de Livos")
  print("2 - Busca de Livros")
  print("3 - Verificação de acervo")
  print("4 - Emprestimos")
  print("5 - Devolução")
  print("6 - Cadastro de Leitores")
  print("7 - Sair")
  print()
  global opcao # (variavel global para ser usada em outros blocos do código)
  opcao = input(" ")
  print()


# BLOCO 2
# Implementando a primeira opção do Menu principal

# Inserindo livros do arcevo
def cadastro_livros():

    # 1. Definindo os textos que aparecerão antes do input
    textos = [
        'Titulo do Livro:',
        'Subtítulo:',
        'Autores:',
        'Editora:',
        'Edição:',
        'Ano de Publicação:',
        'Local de Publicação:',
        'ISBN:',
        'Numero de Páginas:',
        'Idioma:',
        'Código de Barras:',
        'Data de entrada:',
        'Forma de obtenção:',
        'Valor da Obra:',
        'Número de Exemplares:',
        'Condição do Exemplar:',
        'Assuntos:',
        'Genero:',
        'Categoria:',
        'Sinopse:',
        'Status do Item:',
        'Histórico de Empréstimos:']

    # 2. Descobrindo o comprimento do maior texto
    maior_comprimento = max(len(t) for t in textos)

    # 3. Recebendo as informações com os textos alinhados
    print("=" *74)
    print('Por favor insira dos dados a seguir!')
    print ("=" *74)
    print()
    titulo = input(textos[0].ljust(maior_comprimento + 1))
    Subtítulo = input(textos[1].ljust(maior_comprimento + 1))
    Autores = input(textos[2].ljust(maior_comprimento + 1))
    Editora = input(textos[3].ljust(maior_comprimento + 1))
    Edição = input(textos[4].ljust(maior_comprimento + 1))
    Ano_de_Publicacao = input(textos[5].ljust(maior_comprimento + 1))
    Local_de_Publicacao = input(textos[6].ljust(maior_comprimento + 1))
    ISBN = input(textos[7].ljust(maior_comprimento + 1))
    N_de_Páginas = input(textos[8].ljust(maior_comprimento + 1))
    Idioma = input(textos[9].ljust(maior_comprimento + 1))
    Código_de_Barras = input(textos[10].ljust(maior_comprimento + 1))
    Data_de_entrada = input(textos[11].ljust(maior_comprimento + 1))
    forma_de_obtenção = input(textos[12].ljust(maior_comprimento + 1))
    Valor_Da_Obra = input(textos[13].ljust(maior_comprimento + 1))
    Número_de_Exemplares = input(textos[14].ljust(maior_comprimento + 1))
    Condição_do_Exemplar = input(textos[15].ljust(maior_comprimento + 1))
    Assuntos = input(textos[16].ljust(maior_comprimento + 1))
    Genero = input(textos[17].ljust(maior_comprimento + 1))
    Categoria = input(textos[18].ljust(maior_comprimento + 1))
    Sinopse = input(textos[19].ljust(maior_comprimento + 1))
    Status_do_Item = input(textos[20].ljust(maior_comprimento + 1))
    Histórico_de_Empréstimos = input(textos[21].ljust(maior_comprimento + 1))

    # Organizando as 22 informações na ordem do cadastro
    dados_livro = [
        titulo,
        Subtítulo,
        Autores,
        Editora,
        Edição,
        Ano_de_Publicacao,
        Local_de_Publicacao,
        ISBN,
        N_de_Páginas,
        Idioma,
        Código_de_Barras,
        Data_de_entrada,
        forma_de_obtenção,
        Valor_Da_Obra,
        Número_de_Exemplares,
        Condição_do_Exemplar,
        Assuntos,
        Genero,
        Categoria,
        Sinopse,
        Status_do_Item,
        Histórico_de_Empréstimos]

    # Abrindo a planilha de livros já carregada
    planilha = cadastrolivros.active

    # Encontrando a próxima linha vazia pela coluna B
    proxima_linha = 4

    while planilha.cell(row=proxima_linha, column=2).value is not None:
        proxima_linha += 1

    # Inserindo os dados a partir da coluna B
    for coluna, valor in enumerate(dados_livro, start=2):
        celula = planilha.cell(row=proxima_linha,column=coluna)
        celula.value = valor

        # Copiando a formatação da linha anterior
        if proxima_linha > 2:
            celula_anterior = planilha.cell(row=proxima_linha - 1, column=coluna)

            if celula_anterior.has_style:
                from copy import copy
                celula._style = copy(celula_anterior._style)

            celula.number_format = celula_anterior.number_format

    # Salvando no arquivo Excel existente
    cadastrolivros.save('Tabelas/Cadastro de Livros.xlsx')
    print()
    print("=" *74)
    print("Livro cadastrado e salvo na planilha com sucesso!")
    print("=" *74)
    print()
    
menu_principal()
if opcao == "1":
   cadastro_livros()
       

# BLOCO 3
#Implementando a segunda opção do Menu principal

def busca_livros():
    titulo_busca = input(str("Digite o título do livro que deseja buscar: "))
    print()
    planilha = cadastrolivros.active
    for linha in range(1, planilha.max_row + 1):
        titulo_planilha = planilha.cell(row=linha, column=2).value
        if titulo_planilha is not None and titulo_planilha.strip().lower() == titulo_busca.strip().lower():
           print ("="* 74)
           print("******************* Livro encontrado!!!*******************")
           print("=" * 74)
           print()
           print()

           for coluna in range(1, 24): 
                campo = planilha.cell(row=1, column=coluna).value
                valor = planilha.cell(row=linha, column=coluna).value
                
                print(f"{str(campo):<35} = {valor}")
                print("-" * 74)
           break
    else:
        print("=" * 74)
        print(" **************Livro não encontrado na base do Arcevo.***************")
        print("=" * 74)
        print()
        print()


if opcao == "2":
   busca_livros()

# BLOCO 4
#Implementando a terceira opção do Menu principal

def verificacao_arcevo():
    import pandas as pd
    cadastrolivros = pd.read_excel('Tabelas/Cadastro de Livros.xlsx')
    print("=" * 120)
    print("                                         ARCEVO DE LIVROS CADASTRADOS NA BIBLIOTECA MUNICIPAL DE CAMOCIM")
    print("=" * 120)
    print()
    
    # PRIMEIRA TABELA - 5 COLUNAS
    print("-" * 120)
    print(cadastrolivros.iloc[:, 1:6].to_string(index=False, justify='rihght', col_space=40)) 
    print("-" * 120)

    # SEGUNDA TABELA - 5 COLUNAS
    print()
    print("-" * 120)
    print(cadastrolivros.iloc[:, 6:12].to_string(index=False, justify='rihght', col_space=40))
    print("-"* 120)

    #TERCEIRA TABELA - 5 COLUNAS
    print()
    print("-"* 120)
    print(cadastrolivros.iloc[:,12:18].to_string(index=False, justify='rihght', col_space=40))
    print("-" * 120)
    
    #QUARTA  TABELA - 5 COLUNAS
    print()
    print(cadastrolivros.iloc[:, 18:23].to_string(index=False, justify='rihght', col_space=40))
    print("=" * 120)
    print()
    print("=" * 120)
    print('                                                  Lista do arcevo foi impressa')
    print("=" * 120)
    print()
    

if opcao == "3":
    verificacao_arcevo()


# BLOCO 5
# Implementado a Quarta opção do Menu principal (4 - Empréstimo)

def emprestimo_livros():

    print("=" * 74)
    print("                         EMPRÉSTIMO DE LIVROS")
    print("=" * 74)
    print()

    chave = input(" Digite o CPF do Leitor: ")
    print()

    planilha = cadastroleitores.active

    for linha in range(2, planilha.max_row + 1):

        Id_cliente = planilha.cell(row=linha, column=2).value

        if Id_cliente is not None and str(Id_cliente).strip().replace(".", "").replace("-", "").lstrip("0") == chave.strip().replace(".", "").replace("-", "").lstrip("0"):

            print("=" * 74)
            print("*************** Cliente Cadastrado!!! ***************")
            print("=" * 74)

            nome = planilha.cell(row=linha, column=3).value

            print()
            print(f" Nome do cliente: {nome.upper()}")
            print()

            celular = input(" Digite o celular do cliente: ")
            print()

            titulo = input(" Digite o título do livro: ")
            print()

            planilha_livros = cadastrolivros.active
            planilha_emprestimo = controleemprestimo.active

            # Procurar o livro pelo título na coluna 2

            for linha_livro in range(2, planilha_livros.max_row + 1):

                if str(planilha_livros.cell(row=linha_livro, column=2).value).strip().lower() == titulo.strip().lower():

                    # Verificar disponibilidade na coluna 22

                    situacao_livro = planilha_livros.cell(row=linha_livro, column=22).value

                    if situacao_livro is not None and str(situacao_livro).strip().lower() == "disponivel":

                        linha_emprestimo = 2

                        while planilha_emprestimo.cell(row=linha_emprestimo, column=1).value is not None:
                            linha_emprestimo += 1

                        # Preencher os dados do empréstimo

                        planilha_emprestimo.cell(row=linha_emprestimo, column=1).value = str(Id_cliente).zfill(11)
                        planilha_emprestimo.cell(row=linha_emprestimo, column=2).value = nome
                        planilha_emprestimo.cell(row=linha_emprestimo, column=3).value = celular
                        planilha_emprestimo.cell(row=linha_emprestimo, column=4).value = titulo
                        planilha_emprestimo.cell(row=linha_emprestimo, column=6).value = date.today()

                        # Preencher a data prevista de devolução

                        planilha_emprestimo.cell(row=linha_emprestimo, column=7).value = f"=F{linha_emprestimo}+31"
                        planilha_emprestimo.cell(row=linha_emprestimo, column=7).number_format = "DD/MM/YYYY"

                        # Alinhar os dados à direita

                        for coluna in range(1, 11):
                            planilha_emprestimo.cell(row=linha_emprestimo, column=coluna).alignment = Alignment(horizontal="right")

                        # Salvar a planilha de empréstimos

                        controleemprestimo.save("Tabelas/Controle de Emprestimo.xlsx")

                        # Atualizar a situação do livro

                        planilha_livros.cell(row=linha_livro, column=22).value = "Indisponível"
                        planilha_livros.cell(row=linha_livro, column=23).value = date.today()

                        cadastrolivros.save("Tabelas/Cadastro de Livros.xlsx")

                        print()
                        print("=" * 74)
                        print("           Empréstimo cadastrado com sucesso!")
                        print("=" * 74)
                        print()

                    else:

                        print()
                        print("=" * 74)
                        print(" LIVRO NÃO DISPONÍVEL PARA EMPRÉSTIMO!")
                        print("=" * 74)
                        print()

                    break

            else:

                print()
                print("=" * 74)
                print(" LIVRO NÃO ENCONTRADO NO ACERVO DA BIBLIOTECA!")
                print("=" * 74)
                print()

            break


if opcao == "4":
    emprestimo_livros()

#BLOCO 6 (5 - DEVOLUÇÃO DE EMPRESTIMOS DE LIVROS)

def devolucao_emprestimos():

    print("=" * 74)
    print("                         DEVOLUÇÃO DE LIVROS")
    print("=" * 74)
    print()

    chave = input(str(" Digite o CPF do Leitor: "))
    print()

    planilha = cadastroleitores.active

    for linha in range(2, planilha.max_row + 1):
        Id_cliente = planilha.cell(row=linha, column=2).value

        if (Id_cliente is not None and str(Id_cliente).strip().replace(".", "").replace("-", "").lstrip("0")
           == chave.strip().replace(".", "").replace("-", "").lstrip("0")):

           print("=" * 74)
           print("*************** Cliente Cadastrado!!! ***************")
           print("=" * 74)

           nome = planilha.cell(row=linha, column=3).value

           print()
           print(f" Nome do cliente: {nome.upper()}")
           print()

           celular = input(" Digite o celular do leitor: ")
           print()

           titulo = input(" Digite o título do livro: ")

            # Acessar a planilha de empréstimos
           planilha_emprestimo = controleemprestimo.active

           for linha_emprestimo in range(2, planilha_emprestimo.max_row + 1):

               # Condição dupla: CPF e título do livro
               titulo_cadastrado = planilha_emprestimo.cell(row=linha_emprestimo, column=4).value
               Id_cliente = planilha_emprestimo.cell(row=linha_emprestimo, column=1).value

               if (Id_cliente is not None and str(Id_cliente).strip().replace(".", "").replace("-", "").zfill(11)
                  == chave.strip().replace(".", "").replace("-", "").zfill(11) and titulo_cadastrado is not None
                  and str(titulo_cadastrado).strip().lower() == titulo.strip().lower()):


                  # Atualizar a data efetiva de devolução
                  planilha_emprestimo.cell(row=linha_emprestimo, column=8).value = date.today()
                  planilha_emprestimo.cell(row=linha_emprestimo, column=8).number_format = "DD/MM/YYYY"
                                                     
                   
                  # Atualizar situação do livro no Cadastro de Livros
                  planilha_livros = cadastrolivros.active

                  for linha_livro in range(2, planilha_livros.max_row + 1):

                      titulo_livro = planilha_livros.cell(row=linha_livro, column=2).value

                      if titulo_livro is not None and str(titulo_livro).strip().lower() == str(titulo).strip().lower():
                            
                        planilha_livros.cell(row=linha_livro, column=22).value = "disponivel"
                        planilha_livros.cell(row=linha_livro, column=23).value = date.today()
                        planilha_livros.cell(row=linha_livro, column=23).number_format = "DD/MM/YYYY"
                        break
                  else:
                     print()
                     print(" LIVRO NÃO ENCONTRADO NO CADASTRO DE LIVROS!")
                     print()
                     return        
                  
                  cadastrolivros.save("Tabelas/Cadastro de Livros.xlsx")
                  controleemprestimo.save("Tabelas/Controle de Emprestimo.xlsx")
                  

                  print()
                  print("=" * 74)
                  print(" DEVOLUÇÃO REGISTRADA COM SUCESSO!")
                  print("=" * 74)
                  print()
                  return 
           
if opcao == "5":
    devolucao_emprestimos()
                
# BLOCO 7 (CADASTRO DO LEITOR)

def cadastro_leitor():

    # 1. Definindo os textos que aparecerão antes do input
    textos = [
        'CPF:',
        'Nome:',
        'Endereço:',
        'Bairro:',
        'Cidade:',
        'Celular:']

    # 2. Descobrindo o comprimento do maior texto
    maior_comprimento = max(len(t) for t in textos)

    # 3. Recebendo as informações com os textos alinhados
    print("=" *74)
    print('Por favor insira dos dados a seguir!')
    print ("=" *74)
    print()

    CPF = input(textos[0].ljust(maior_comprimento + 1))
    Nome = input(textos[1].ljust(maior_comprimento + 1))
    Endereço = input(textos[2].ljust(maior_comprimento + 1))
    Bairro = input(textos[3].ljust(maior_comprimento + 1))
    Cidade = input(textos[4].ljust(maior_comprimento + 1))
    celular = input(textos[5].ljust(maior_comprimento + 1))
    
    # Organizando as 22 informações na ordem do cadastro
    dados_leitor = [CPF,Nome, Endereço, Bairro, Cidade, celular]

    # Abrindo a planilha de livros já carregada
    planilha = cadastroleitores.active

    # Encontrando a próxima linha vazia pela coluna B
    proxima_linha = 2

    
    while planilha.cell(row=proxima_linha, column=2).value is not None:
        proxima_linha += 1

    # Inserindo os dados a partir da coluna B
    for coluna, valor in enumerate(dados_leitor, start=2):
        celula = planilha.cell(row=proxima_linha,column=coluna)
        celula.value = valor

        # Copiando a formatação da linha anterior
        if proxima_linha > 2:
            celula_anterior = planilha.cell(row=proxima_linha - 1, column=coluna)

            if celula_anterior.has_style:
                from copy import copy
                celula._style = copy(celula_anterior._style)

            celula.number_format = celula_anterior.number_format

    # Salvando no arquivo Excel existente
    cadastroleitores.save('Tabelas/Cadastro de Leitores.xlsx')
    print()
    print("=" *74)
    print("Leitor Cadastrado e salvo na base com sucesso!")
    print("=" *74)
    print()
    
menu_principal()
if opcao == "6":
    cadastro_leitor()

# BLOCO 8 -  FINAL ( Retorno ao Menu incial)
# INICIALIZAÇÃO DO SISTEMA
while True:

    menu_principal()

    if opcao == "1":
        cadastro_livros()

    if opcao == "2":
        busca_livros()

    if opcao == "3":
        verificacao_arcevo()

    if opcao == "4":
        emprestimo_livros()

    if opcao == "5":
        devolucao_emprestimos()

    if opcao == "6":
        cadastro_leitor()

    if opcao == "7":
        print("Sistema encerrado!")
        break

