import openpyxl
from random import sample

# def numero_aleatorio(tamanho_list, inicial, final):
#     y = sorted(sample(range(inicial, final), tamanho_list))
#     return y
#
#
# tam = int(input('Digite o tamenho da lista aleatoria: '))
# ini = int(input('Digite o numero inicial do intervalo da lista aleatoria: '))
# fim = int(input('Digite o numero final do intervalo da lista aleatoria: '))
# lista_aleatoria = numero_aleatorio(tam, ini, fim)
# print(lista_aleatoria)
Coluna = input('Digite a letra que representa a coluna do excel desejada: ')
nome = 0
lista_final = list()
lista_final_max = list()
lista_final_min = list()
# for c in lista_aleatoria:
for c in range(1, 542):
    arquivo = openpyxl.load_workbook(filename=f"C:\\Users\\pedro\\OneDrive - Office 365\\UFU\\Imagens médicas\\Planilhas\\planilha_com_COVID\\1_informacoesdicom  ({c}).xlsx")
    planilha = arquivo['Planilha1']
    # print(f'Arquivo{c}:  {planilha["v2"].value}')
    lista = list()
    a = 1
    print(c)
    while True:
        a += 1
        nome = planilha[f'{Coluna}1'].value
        val_celula = planilha[f'{Coluna}{a}'].value
        if val_celula == None:
            break
        elif type(val_celula) == str:
            z = val_celula.split(sep='Y')
            if z[0].isnumeric():
                val_celula = int(z[0])
                lista.append(val_celula)
                if val_celula <= 17:
                    print(f'{a} = {val_celula}')
        elif type(val_celula) != str:
            lista.append(val_celula)
    maximo = max(lista)
    minimo = min(lista)
    lista_final_max.append(maximo)
    lista_final_min.append(minimo)
    soma = sum(lista)
    qntd = len(lista)
    if qntd != 0:
        med = soma/qntd
        lista_final.append(med)
soma = sum(lista_final)
qntd = len(lista_final)
print(qntd)
med = soma/qntd
print(f'Media final do {nome}: {med:.3f}')
print(f'Maior número do {nome}: {max(lista_final_max)}')
print(f'Menor valor do {nome}: {min(lista_final_min)}')

