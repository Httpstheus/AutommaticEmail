import win32com.client as win32 

outlook = win32.Dispatch('outlook.application') # É necessário ter o Outlook baixado e vínculado a uma conta, pois ele será utilizado para enviar as mensagens

mail = outlook.CreateItem(0)
mail.To =  '' #Inserir o Destinátario
mail.Subject = '' # Assunto do E-mail

# As aspas triplas ''' indica que todo o conteúdo dentro dela é um HTML, pode usar as tags e styles
mail.HTMLBody = '''
    <p> Seu conteúdo HTML aqui </p>
'''

anexo = r"C:\caminho\do\seu\arquivo\arquivo.xls" # Caso queira enviar anexo, é só substituir o caminho do arquivo, pode ser qualquer extensão

mail.Attachments.Add(anexo)
mail.Send()
print('Email Enviado')