#!usr/bin/python3
import firebirdsql as fb
import xlrd
import pandas as pd
import socket
from contextlib import closing
from datetime import datetime

socket.getaddrinfo('localhost', 8080)
con = fb.connect(host='localhost', database='caminho do banco',user='usuario', password='senha')
NFCe = pd.read_excel('caminho do excel')
log = open('caminho no txt para saida','w')

cur = con.cursor()
#variaveis que armazenam as colunas do excel
excel_Num = NFCe['Numero']
excel_Situacao = NFCe['Situacao']
excel_data = NFCe['Data_Emissao']
excel_chave = NFCe['Chave_Acesso']
excel_valor = NFCe['Valor N.F.']
excel_chave_edit = []
listNFCE = excel_Num.tolist()
cont = 0
cupons_existentes = []
notas_nulas = []
CHV_NFCE = []
VALOR = []
DESCTO = []
cancelados = 0
nao_cancelados = 0
dados_corretos = 0
valor_real = 0
param_cupom_cancelado = 'S'
mensagem_cancelados = ''
mensagem_dado_errado = ''
#loop para passar pelo excel
for i in range(len(listNFCE)):
    mensagem = "Cupom: "+str(listNFCE[i])
    #query para conferir se o numero do cupom no banco corresponde ao do excel
    cur.execute('SELECT COUNT(1) FROM PEDIDOMT WHERE NUM_NFCE = (?)',(excel_Num[i],))

    cont = cont + 1
    #cupons_existentes[i] += listNFCE[i]
    if cur.fetchone()[0]:
        mensagem += " existe no excel e banco"
    else:
        mensagem += "nao existe no banco"
        log.write(mensagem)
        log.write('\n')
    #caso tenha cancelamento no excel, esta query e executada para verificar se foi cancelado no banco
    if excel_Situacao[i] == "Cancelamento de NF-e homologado":
        cur.execute('SELECT COUNT(1) FROM PEDIDOMT WHERE NUM_NFCE = (?) AND CUPOM_CANCELADO = (?)',(excel_Num[i],param_cupom_cancelado,))
        if cur.fetchone()[0]:
            cancelados += 1
        else:
            nao_cancelados += 1
            mensagem_cancelados += 'cupom nao cancelado no banco : ' + str(excel_Num[i])
            log.write(mensagem_cancelados)
            log.write('\n')
    elif excel_Situacao[i] != "Cancelamento de NF-e homologado":
    #query para checar se os dados no excel corresponde ao do banco
        excel_chave[i] = excel_chave[i][3:]
        cur.execute('SELECT CHV_NFCE,VALOR-DESCTO TOTAL FROM PEDIDOMT WHERE CHV_NFCE = (?) AND VALOR-DESCTO = (?)',(excel_chave[i],excel_valor[i],))
        readfromdb = cur.fetchone()
        if readfromdb is not None:
                CHV_NFCE,TOTAL = list(readfromdb)
        else:
            mensagem_dado_errado = 'cupom possui algum dado errado :' + str(excel_Num[i]) + ', valor no excel: ' + str(excel_valor[i]) + ', valor no banco: ' + str(TOTAL) + ', '
            log.write(mensagem_dado_errado)
            log.write('\n')

log.write('passou')
log.close()
