from openpyxl import Workbook, load_workbook
import numpy as np
import matplotlib.pyplot as plt


#importando a planilha de dados
planilha = load_workbook("Dados.xlsx")
aba_ativa = planilha.active

#criando uma tabela para todos os dados dentro do python
dados = []
for celula in aba_ativa["B"]:
    if type(celula.value) == int:
        dados.append(celula.value)
#separando os dados de acordo com os meses
mes1 = []
for celula in aba_ativa["A"]:
    if str(celula.value)[3] == '2':
        linha = celula.row
        mes1.append(aba_ativa[f'B{linha}'].value)

mes2 = []
for celula in aba_ativa["A"]:
    if str(celula.value)[3] == '3':
        linha = celula.row
        mes2.append(aba_ativa[f'B{linha}'].value)

#calculando a media de cada mes
media_mes1 = np.mean(mes1)
media_mes2 = np.mean(mes2)

#usando o metodo da média móvel com suavização exponencial (mmse)
variacao_de_venda = (media_mes2 - media_mes1) / media_mes1
demanda_media = np.mean(dados)
mmse = (variacao_de_venda * demanda_media) + (1 - variacao_de_venda) * (945)
#fazendo a previsão para os proximos 5 dias
previsao = []

for c in range(0,5):   
    mmse = (variacao_de_venda * demanda_media) + (1 - variacao_de_venda) * (mmse)
    previsao.append(int(mmse))

#imprimindo o resultado
    dia = 21
for c in range(0,5):
    print("a previsão para o dia {} é de {}".format(dia, previsao[c]))
    dia+=1

#Crianco um gráfico para melhor visualização do dados

dias = ['21', '22', '23', '24', '25']
plt.xlabel("Dias")
plt.ylabel("Vendas")
plt.plot(dias,previsao)
plt.show()







