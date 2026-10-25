# Instalação e importação das Bibliotecas
#pip install xlsxwriter
#pip install openpyxl
#import xlsxwriter

#import openpyxl
from openpyxl import load_workbook
from copy import copy

#importação dos arquivos (Pesquisa e aulas assistida como montar drive)
cadastroleitores = load_workbook('Tabelas/Cadastro de Leitores.xlsx')
cadastrolivros = load_workbook('Tabelas/Cadastro de Livros.xlsx')

# Iniciando o sistema com apresentação do Menu Princial
def menu_principal():
# Mensagem iniciala
  print("=" *74)
  print(" Seja Bem-vindos ao Sistema de Gestão da Biblioteca Municipal de Camocim!")
  print("=" *74)
  print()

#Apresentação do Menu Inicail
  print ("Por favor selecione uma das opções a baixo:")
  print()
  print()
  print("1 - Cadastro de Livos")
  print("2 - Busca de Livros")
  print("3 - Verificação de acervo")
  print("4 - Emprestimos")
  print("5 - Devolução")
  print("6 - Cadastro de Leitores")
  print("7 - Sair")
  print()
  opcao = input(" ")
  print()



  if opcao == "1":
     cadastro_livros()




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
    
    pass

menu_principal()