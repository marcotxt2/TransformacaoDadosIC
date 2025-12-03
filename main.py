import tabula
import pandas as pd
import zipfile

Arquivo_Zip_Nome = "Teste_Marco.zip"
Arquivo_Csv = "Anexo_I.csv"
Arquivo_Pdf = "anexoI/Anexo_I.pdf"
Minimo_Linhas = 10 #Defini 10 pois identifiquei que a menor tabela do pdf tem 10 linhas apenas.

#Com a biblioteca tabula, o programa identifica todas as tabelas do PDF. (lattice para identificar as linhas com bordas definidas)
Lista_Tabelas = tabula.read_pdf(Arquivo_Pdf, pages= "all",encoding='utf-8', lattice = True)

#For loop para analisar se a tabela tem menor ou igual a 10 linhas (garantindo assim que seja uma tabela válida do PDF)
Tabelas_Reais = []
for tabela in Lista_Tabelas:
    if len(tabela) >= Minimo_Linhas:
        Tabelas_Reais.append(tabela)
print("Tabelas a serem extraidas: ", len(Tabelas_Reais))

if not Tabelas_Reais:
    print ("Nenhuma tabela válida encontrada.")

if len(Tabelas_Reais) > 0:
    #Empilha as tabelas
    Tabela_Completa = pd.concat(Tabelas_Reais, ignore_index=True) #ignore_index deixa de resetar o index para cada linha da tabela.

    #Identifica a coluna "OD" e "AMB", faz o replace de cada para o inserido no rodapé das páginas.
    #Obs.: Utilizado regex para não ter erro do tabula ter identificado os elementos com espaço vazio por exemplo " OD"...
    if "OD" in Tabela_Completa.columns:
        Tabela_Completa["OD"] = Tabela_Completa["OD"].str.replace("OD","Seg. Odontológica", regex = True)
    if "AMB" in Tabela_Completa.columns:
        Tabela_Completa["AMB"] = Tabela_Completa["AMB"].str.replace("AMB","Seg. Ambulatorial", regex = True)

    print("Trocando os elementos OD e AMB por Seg. Odontológica e Seg. Ambulatorial respectivamente")

    #Transforma para CSV a tabela completa.
    try:
        Tabela_Completa.to_csv(Arquivo_Csv, index=False, encoding='utf-8-sig')
        print("CSV Criado: ", Arquivo_Csv)

        #Pega o arquivo "Anexo_I.csv" e guarda dentro do arquivo zip "AnexoI_CSV.zip".
        with zipfile.ZipFile(Arquivo_Zip_Nome, 'w') as Arquivo_Zip:
            Arquivo_Zip.write(Arquivo_Csv)
        print("ZIP criado: ", Arquivo_Zip_Nome)

    except Exception as e:
        print("Erro ao salvar arquivos: ", e)