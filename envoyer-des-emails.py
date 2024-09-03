import win32com.client as win32

# creer une intgration avec l'outlook
outlook = win32.Dispatch('outlook.application')

# creer un e-mail
email = outlook.CreateItem(0)


# configuratio des infos du e-mail
email.To = "marcosomarcal@yahoo.com.br"
email.Subject = "E-mail automático do Python"
email.HTMLBody = f"""
<p>Ce e-mail a été envoyé avec le python</p>

<p>codePython</p>
"""

# anexo = "C://Users/marc/Downloads/arquivo.xlsx"
# email.Attachments.Add(anexo)

email.Send()
print("Email envoyé")
