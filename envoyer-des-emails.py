import win32com.client as win32

# criar a integração com o outlook
outlook = win32.Dispatch('outlook.application')

# criar um email
email = outlook.CreateItem(0)


# configurar as informações do seu e-mail
email.To = "marcosomarcal@yahoo.com.br"
email.Subject = "E-mail automático do Python"
email.HTMLBody = f"""
<p>teste de e-mail sendo enviado com python</p>

<p>Código Python</p>
"""

# anexo = "C://Users/marc/Downloads/arquivo.xlsx"
# email.Attachments.Add(anexo)

email.Send()
print("Email Enviado")
