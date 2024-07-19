from openpyxl import load_workbook
import os

import win32com.client as win32

outlook = win32.Dispatch('Outlook.application')

nomeCaminhoArquivo = "E:\\projRPA\\rpaChrome\\listEmail.xlsx"
planilha_aberta = load_workbook(filename=nomeCaminhoArquivo)

planilha_selecionada = planilha_aberta['Planilha1']

for linha in range(2, len(planilha_selecionada['A']) + 1):


    nome = planilha_selecionada['A%s' % linha].value
    nomeCompleto = planilha_selecionada['B%s' % linha].value
    endmail = planilha_selecionada['C%s' % linha].value


    emailOutlook = outlook.CreateItem(0)

    emailOutlook.to = endmail
    emailOutlook.Subject = "Envio de RPA -> Teste" + nomeCompleto

    emailOutlook.HTMLBody = f""""
    <p>Boa Noite <b>{nome}</b>.</p>
    <p>Teste de envio de e-mail com anexo.</p>
    <p>Atenciosamente.</p>
    
    """

    anexoMail = "E:\\projRPA\\listEmail.xlsx" + nomeCompleto + ".xlsx"
    emailOutlook.Attachments.Add(anexoMail)
    #email.Attachments.Add(anexoMail)


    emailOutlook.Send()

    print('E-mail enviado com sucesso')