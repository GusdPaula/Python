import win32com.client as win32
from openpyxl import load_workbook
import os


outlook = win32.Dispatch('outlook.application')


caminhoSistema = 'C:\\Users\\gustavo.paula\\OneDrive - SCLA\\Desktop\\Python +excel como automatizar processos\\ListaEmails.xlsx'
planilhalista = load_workbook(filename = caminhoSistema)

sheetLista = planilhalista['Dados']

#CRIANDO RANGE PARA ENVIO
for i in range(2, len(sheetLista['B']) +1):

    emailOutlook = outlook.CreateItem(0)

    nome = sheetLista['A%s' % i].value
    email = sheetLista['B%s' % i].value

    emailOutlook.To = email
    emailOutlook.Subject = "E-mail com Python" + nome
    #colocando f é possível colocar variáveis
    emailOutlook.HTMLBody = f"""
    <p>Boa noite, {nome}!<p/>
    <p>Apenas avisando que apendemos a criar um e-mail no python e enviá-lo automaticamente<p/>
    <p>Atenciosamente, Gustavo.<p/>
    <p><img src = "C:\\Users\\gustavo.paula\\OneDrive - SCLA\\Pictures\\ass.PNG"> </p>
    """
    #ANEXANDO ARQUIVOS
    anexoexcel = "C:\\Users\\gustavo.paula\\OneDrive - SCLA\\Desktop\\Python +excel como automatizar processos\\ListaEmails.xlsx"
    emailOutlook.Attachments.Add(anexoexcel)
    
    #ENVIO DOS E-MAILS
    emailOutlook.Send() #.Send envia e .Save salva
    
    #CONFIRMAÇÃO DE ENVIO
    print("E-mail enviado com sucesso para", nome, "!")
