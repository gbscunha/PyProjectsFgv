import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import locale


def pegando_sheet(a, tamanho, dados, shet):
    for i in range(1, tamanho+1):
        erro = False
        try:
            arquivo_exel2 = pd.read_excel(a, sheet_name=i)
            aux1 = arquivo_exel2.columns[0][0].upper()
            aux1 = aux1+arquivo_exel2.columns[0][1::]
            arquivo_exel2 = arquivo_exel2.values
        except:
            print('\nErro ao encontrar o nome do ct!')
            print("\nO nome da Sheet cujo deu erro foi: ")
            try:
                print(shet[i])
            except:
                print("Erro ao indentificar a shet")

        lista = [0]*2
        aux3 = []
        for k in range(4, 20, 5):
            for j in range(4, 6):
                try:
                    pega = arquivo_exel2[k][j]
                    try:
                        aux = round(pega*100)
                        lista[j-4] = aux
                    except:
                        erro = True
                except:
                    print("Erro ao localizar as porcentagens")
                    erro = True
                    try:
                        print(shet[i])
                    except:
                        print("Erro ao indentificar a shet")

            aux3.append(lista)
            lista = [0]*2
        if (erro == False):
            amt = []
            for h in range(3, 20, 5):
                try:
                    am = arquivo_exel2[h][6]
                    am = int(am)
                except:
                    print("erro na conversão ou erro na procura da amostra")
                amt.append(am)
            print(amt)
            if (amt[0] >= amt[1] and amt[0] >= amt[2] and amt[0] >= amt[3]):
                maior = amt[0]
            elif (amt[1] >= amt[2] and amt[1] >= amt[3]):
                maior = amt[1]
            elif (amt[2] >= amt[3]):
                maior = amt[2]
            else:
                maior = amt[3]
            dicio = {'CT': aux1, 'local': aux3[0], 'ambiente': aux3[1],
                     'qualidade': aux3[2], 'etica': aux3[3], 'amostra': maior}
            dados.append(dicio)
    return (dados)


def colocar_emordem(dados):
    locale.setlocale(locale.LC_ALL, 'pt_br.utf-8')
    dados = sorted(dados, key=lambda dados: (locale.strxfrm(dados['CT'])))
    return (dados)


def calculando_porcentagem(dados, tamanho):
    cts_problema = []
    for i in range(len(dados)):
        local = dados[i]['local'][0]+dados[i]['local'][1]
        ambiente = dados[i]['ambiente'][0]+dados[i]['ambiente'][1]
        quali = dados[i]['qualidade'][0]+dados[i]['qualidade'][1]
        etica = dados[i]['etica'][0]+dados[i]['etica'][1]
        novo_dados = {'CT': dados[i]['CT'], 'Localização': 'Regular: ' + str(dados[i]['local'][0])+'  e  '+'Ruim: '+str(dados[i]['local'][1]), 'soma_L': str(local),
                      'Ambiente': 'Regular: ' + str(dados[i]['ambiente'][0])+'  e  '+'Ruim: ' + str(dados[i]['ambiente'][1]), 'soma_A': str(ambiente),
                      'Qualidade': 'Regular: ' + str(dados[i]['qualidade'][0])+'  e  '+'Ruim: ' + str(dados[i]['qualidade'][1]), 'soma_Q': str(quali),
                      'Ética': 'Regular: ' + str(dados[i]['etica'][0])+'  e  '+'Ruim: ' + str(dados[i]['etica'][1]), 'soma_E': str(etica), 'Amostra': str(dados[i]['amostra'])}
        if ((dados[i]['local'][0]+dados[i]['local'][1]) >= 10):
            cts_problema.append(novo_dados)
        elif ((dados[i]['ambiente'][0]+dados[i]['ambiente'][1]) >= 10):
            cts_problema.append(novo_dados)
        elif ((dados[i]['qualidade'][0]+dados[i]['qualidade'][1]) >= 10):
            cts_problema.append(novo_dados)
        elif ((dados[i]['etica'][0]+dados[i]['etica'][1]) >= 10):
            cts_problema.append(novo_dados)
        dados[i] = {'CT': dados[i]['CT'], 'Localização': 'Regular: ' + str(dados[i]['local'][0])+'  e  '+'Ruim: '+str(dados[i]['local'][1]), 'soma_L': str(local),
                    'Ambiente': 'Regular: ' + str(dados[i]['ambiente'][0])+'  e  '+'Ruim: ' + str(dados[i]['ambiente'][1]), 'soma_A': str(ambiente),
                    'Qualidade': 'Regular: ' + str(dados[i]['qualidade'][0])+'  e  '+'Ruim: ' + str(dados[i]['qualidade'][1]), 'soma_Q': str(quali),
                    'Ética': 'Regular: ' + str(dados[i]['etica'][0])+'  e  '+'Ruim: ' + str(dados[i]['etica'][1]), 'soma_E': str(etica), 'Amostra': str(dados[i]['amostra'])}
    return cts_problema, dados


def criando_arquivo(cts_problema, dados):
    df = pd.DataFrame(cts_problema)
    de = pd.DataFrame(dados)
    try:
        with pd.ExcelWriter("Resultado_Programa.xlsx") as writer:
            de.to_excel(writer, sheet_name="Todos os CTs")  
            df.to_excel(writer, sheet_name="CTs de Não Conformidade")  
    except:
        print("\nErro em colocar os dados no arquivo do arquivo")
        print("\nVerifique se o arquivo está aberto\n")
    return ()


Tk().withdraw()
a = askopenfilename()
arquivo_exel = pd.read_excel(a)
tamanho = pd.ExcelFile(a)
tamanho = tamanho.sheet_names
shet = tamanho
tamanho = len(tamanho)-1
dados = []
arquivo_exel = arquivo_exel.values
print("\nPegando os nomes dos CTs na Planilha...")
print("\nPegando os dados das porcentagens...")
print("\nEsse processo pode demorar um pouco, por favor aguarde...")
dados = pegando_sheet(a, tamanho, dados, shet)
print("\nOrdenando os dados...")
print("\nCalculando a porcentagem...")
dados = colocar_emordem(dados)
cts_problema, dados = calculando_porcentagem(dados, tamanho)
print("\nCriando arquivo...")
criando_arquivo(cts_problema, dados)
print("\nPrograma finalizado!!\n")
print("Por favor cheque o arquivo na pasta!\n")
