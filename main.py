import tabula
import pandas as pd
import zipfile

arquivoZipNome = "AnexoI_CSV.zip"
arquivoCsv = "Anexo_I.csv"
arquivoPdf = "anexoI/Anexo_I.pdf"
minimoLinhas = 10 #Defini 10 pois identifiquei que a menor tabela do pdf tem 10 linhas apenas.

#Com a biblioteca tabula, o programa identifica todas as tabelas do PDF. (lattice para identificar as linhas com bordas definidas)
listaTabelas = tabula.read_pdf(arquivoPdf, pages= "all",encoding='utf-8', lattice = True)

#For loop para analisar se a tabela tem menor ou igual a 10 linhas (garantindo assim que seja uma tabela válida do PDF)
tabelasReais = []
for tabela in listaTabelas:
    if len(tabela) >= minimoLinhas:
        tabelasReais.append(tabela)
print("Tabelas a serem extraidas: ", len(tabelasReais))

if not tabelasReais:
    print ("Nenhuma tabela válida encontrada.")

if len(tabelasReais) > 0:
    #Empilha as tabelas
    tabelaCompleta = pd.concat(tabelasReais, ignore_index=True) #ignore_index deixa de resetar o index para cada linha da tabela.

    #Identifica a coluna "OD" e "AMB", faz o replace de cada para o inserido no rodapé das páginas.
    #Obs.: Utilizado regex para não ter erro do tabula ter identificado os elementos com espaço vazio por exemplo " OD"...
    if "OD" in tabelaCompleta.columns:
        tabelaCompleta["OD"] = tabelaCompleta["OD"].str.replace("OD","Seg. Odontológica", regex = True)
    if "AMB" in tabelaCompleta.columns:
        tabelaCompleta["AMB"] = tabelaCompleta["AMB"].str.replace("AMB","Seg. Ambulatorial", regex = True)

    print("Trocando os elementos OD e AMB por Seg. Odontológica e Seg. Ambulatorial respectivamente")

    #Transforma para CSV a tabela completa.
    try:
        tabelaCompleta.to_csv(arquivoCsv, index=False, encoding='utf-8-sig')
        print("CSV Criado: ", arquivoCsv)

        #Pega o arquivo "Anexo_I.csv" e guarda dentro do arquivo zip "AnexoI_CSV.zip".
        with zipfile.ZipFile(arquivoZipNome, 'w') as arquivoZip:
            arquivoZip.write("Anexo_I.csv")
        print("ZIP criado: ", arquivoZipNome)

    except Exception as e:
        print("Erro ao salvar arquivos: ", e)