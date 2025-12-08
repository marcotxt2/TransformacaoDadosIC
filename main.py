import tabula
import pandas as pd
import zipfile

ARQUIVO_ZIP_NOME = "Teste_Marco.zip"
ARQUIVO_CSV = "Anexo_I.csv"
ARQUIVO_PDF = "anexoI/Anexo_I.pdf"
MINIMO_LINHAS = 10 #Defini 10 pois identifiquei que a menor tabela do pdf tem 10 linhas apenas.

def filtrar_tabelas_reais(minimo_linhas: int, lista_tabelas: list) -> list:
    """
    Filtra as tabelas que possuem um número mínimo de linhas.
    :param minimo_linhas: Pega o valor minimo de linhas para ser definido real.
    :param lista_tabelas: Pega o array das tabelas fornecidos pela tabula.
    :return: Retorna a lista de tabelas reais (válidas).
    """
    tabelas_reais = []
    for tabela in lista_tabelas:
        if len(tabela) >= minimo_linhas:
            tabelas_reais.append(tabela)
    return tabelas_reais

def main():
    #Com a biblioteca tabula, o programa identifica todas as tabelas do PDF. (lattice para identificar as linhas com bordas definidas)
    try:
        print("Iniciando identificação de tabelas no PDF...")
        lista_tabelas = tabula.read_pdf(ARQUIVO_PDF, pages="all", encoding='utf-8', lattice = True)
    except Exception as e:
        print(f"Erro ao ler arquivos: {e} ")
        return

    #For loop para analisar se a tabela tem menor ou igual a 10 linhas (garantindo assim que seja uma tabela válida do PDF)
    tabelas_reais = filtrar_tabelas_reais(MINIMO_LINHAS, lista_tabelas)
    qtd_tabelas_reais = len(tabelas_reais)

    if not qtd_tabelas_reais:
        print ("Nenhuma tabela válida encontrada.")
        return

    print(f"{qtd_tabelas_reais} Tabelas encontradas...")

    #Empilha as tabelas
    tabela_completa = pd.concat(tabelas_reais, ignore_index=True) #ignore_index deixa de resetar o index para cada linha da tabela.

    #Identifica a coluna "OD" e "AMB", faz o replace de cada para o inserido no rodapé das páginas.
    #Obs.: Utilizado regex para não ter erro do tabula ter identificado os elementos com espaço vazio por exemplo " OD"...
    if "OD" in tabela_completa.columns:
        tabela_completa["OD"] = tabela_completa["OD"].str.replace("OD", "Seg. Odontológica", regex = True)
    if "AMB" in tabela_completa.columns:
        tabela_completa["AMB"] = tabela_completa["AMB"].str.replace("AMB", "Seg. Ambulatorial", regex = True)

    print("Trocando os elementos OD e AMB por Seg. Odontológica e Seg. Ambulatorial respectivamente")

    #Transforma para CSV a tabela completa.
    try:
        tabela_completa.to_csv(ARQUIVO_CSV, index=False, encoding='utf-8-sig')
        print("CSV Criado: ", ARQUIVO_CSV)

        #Pega o arquivo "Anexo_I.csv" e guarda dentro do arquivo zip "AnexoI_CSV.zip".
        with zipfile.ZipFile(ARQUIVO_ZIP_NOME, 'w') as arquivo_zip:
            arquivo_zip.write(ARQUIVO_CSV)
        print("ZIP Criado: ", ARQUIVO_ZIP_NOME)

    except Exception as e:
        print(f"Erro ao salvar arquivos: {e}")
main()