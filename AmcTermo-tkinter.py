from tkinter import *
from datetime import date
import openpyxl
import openpyxl.cell._writer

#Recebendo informações e encaminhando para o "Termo" no Excel
def enviarInformacao():
    txtInformacaoEnviada = Label(layout, text="INFORMAÇÕES ENVIADAS", background="#D8BFD8", foreground="#8B008B").place(x=135, y=280)
    #Abrindo Arquivo Excel
    arquivo = openpyxl.load_workbook("TermoDevolucaoEquipamento.xlsx")
    #Abrindo Página do Termo
    planTermo =  arquivo['Termo']
    
    #Inserindo informações recebidas no Excel
    planTermo['B5'] = planTermo['B26'] = "{}/{}/{}".format(dataDia, dataMes, dataAno)
    planTermo['F5'] = planTermo['F26'] = inpLoja.get()
    planTermo['I5'] = planTermo['I26'] = inpCpfAss.get()
    planTermo['J6'] = planTermo['J27'] = inpCpfEnt.get()
    planTermo['C8'] = planTermo['C29'] = inpContato.get()
    planTermo['G4'] = planTermo['G25'] = inpProtocolo.get()
    planTermo['G8'] = planTermo['G29'] = inpCliente.get()
    planTermo['D10'] = planTermo['D31'] = inpEquip.get()
    planTermo['H10'] = planTermo['H31'] = inpCabo.get()
    planTermo['L10'] = planTermo['L31'] = inpControle.get()
    planTermo['A12'] = planTermo['A33'] = inpInfAd.get()

    #Abrindo Página do Gerente
    planGerente = arquivo['Gerente']
    #Inserindo dados recebidos na página do Gerente
    informacoes = ["{}/{}/{}".format(dataDia, dataMes, dataAno), inpCliente.get(), inpCpfAss.get(), inpCpfEnt.get(), inpContato.get(), inpEquip.get(), inpCabo.get(), inpControle.get(), inpProtocolo.get()]
    planGerente.append(informacoes)
    
    arquivo.save("TermoDevolucaoEquipamento.xlsx")
    
#Criando Layout/Janela
layout = Tk()
layout.title("Informações do Termo")
layout.geometry("450x320")
#Data
dataDia = date.today().day
dataMes = date.today().month
dataAno = date.today().year

#Inserindo informações no Layout
Label(layout, text="DADOS - TERMO DE DEVOLUÇÃO DE EQUIPAMENTO", foreground="#8B008B", font=(4)).place(x=10,y=10,width=420, height=30);

Label(layout, text="DATA: {}/{}/{}".format(dataDia, dataMes, dataAno)).place(x=285, y=50);

Label(layout, text="LOJA").place(x=50,y=50);
inpLoja=Entry(layout)
inpLoja.place(x=85, y=50, width=170, height=20)

Label(layout, text="CLIENTE").place(x=50,y=80);
inpCliente=Entry(layout)
inpCliente.place(x=100, y=80, width=280, height=20)

Label(layout, text="CPF ASSINANTE").place(x=50,y=110);
inpCpfAss=Entry(layout)
inpCpfAss.place(x=140,y=110, width=75, height=20)

Label(layout, text="CPF PRESENTE").place(x=223,y=110);  
inpCpfEnt=Entry(layout)
inpCpfEnt.place(x=305,y=110, width=75, height=20)

Label(layout, text="TELEFONE").place(x=50, y=140);
inpContato=Entry(layout)
inpContato.place(x=110, y=140, width=80, height=20)

Label(layout, text="PROTOCOLO").place(x=195, y=140);
inpProtocolo=Entry(layout)
inpProtocolo.place(x=270, y=140, width=110, height=20)

Label(layout, text="EQUIPAMENTO(S)").place(x=50, y=170);
inpEquip=Entry(layout)
inpEquip.place(x=152,y=170,width=30, height=20)

Label(layout, text="CABO(S)").place(x=185, y=170);
inpCabo=Entry(layout)
inpCabo.place(x=237, y=170, width=30, height=20)

Label(layout, text="CONTROLE(S)").place(x=270, y=170);
inpControle=Entry(layout)
inpControle.place(x=350, y=170, width=30, height=20)

Label(layout, text="INFORMAÇÕES ADICIONAIS").place(x=50, y=200);
inpInfAd=Entry(layout)
inpInfAd.place(x=210, y=200, width=170, height=20)

#Enviando informações para "def enviarInformacao()"
btnEnviar = Button(layout, text="Enviar", background="#8B008B", foreground="#ffffff", command=enviarInformacao)
btnEnviar.place(x=160, y=240, width=100, height=30)

layout.mainloop()